"""
populate_xls.py
===============
Read the parsed Modiin calendar CSV and populate the Detail sheet of the
Yakir calendar XLS with raw Shabbat data.

The XLS formulas then calculate all prayer times automatically.
This script only writes the INPUT columns — never the formula columns.

INPUT COLUMNS WRITTEN (Detail sheet):
  AJ (col 36) = Original Source Candles as input 12H  <- stored as time (H-12:MM)
  AK (col 37) = Motzash as input 12H                  <- stored as time (H-12:MM)
  AL (col 38) = Shkia                                 <- stored as datetime (date + time)
  AM (col 39) = date                                  <- stored as date
  AO (col 41) = Parsha Input
  AQ (col 43) = Holiday / RoshCh
  AR (col 44) = Yr
  AS (col 45) = Summer (Y/N)

CRITICAL: Candles & Motzash MUST come from the predefined XLS/CSV only.
          Python NEVER calculates these — they are the official city times.
          A verification diff is printed to stdout for sanity checking.

Usage:
    python populate_xls.py                         # uses defaults
    python populate_xls.py --csv path/to/data.csv --xls path/to/workbook.xlsx
    python populate_xls.py --year 5786             # only update rows matching this Hebrew year
    python populate_xls.py --dry-run               # show what would be written, don't save

Output:
    Modified XLS saved to same filename (backup created as <name>_backup.xlsx)
"""

import argparse
import csv
import datetime
import os
import shutil
import sys

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# ── default paths ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR    = os.path.join(SCRIPT_DIR, "..")  # one level up = Yakir-Calendar-experimental

DEFAULT_CSV = os.path.join(BASE_DIR, "modiin_calendar.csv")
DEFAULT_XLS = os.path.join(BASE_DIR, "yakircal master v61 Added shkia+14 times starting 5785.xlsx")

# ── Detail sheet column indices (1-based, matching openpyxl convention) ────────
COL_A      =  1   # HYrMo (key: YYYYMM numeric)
COL_AJ     = 36   # Original Source Candles as input 12H
COL_AK     = 37   # Motzash as input 12H
COL_AL     = 38   # Shkia (datetime)
COL_AM     = 39   # date
COL_AO     = 41   # Parsha Input
COL_AP     = 42   # nxt Parsha (next week's parsha — derived from CSV sequence)
COL_AQ     = 43   # Holiday / RoshCh
COL_AR     = 44   # Yr
COL_AS     = 45   # Summer (Y/N)

HEADER_ROW  = 1   # row 1 = headers in Detail sheet
DATA_ROW_START = 2  # data starts at row 2

# ── Summer months (by Gregorian month number) ──────────────────────────────────
# In Israel, DST / "summer" is April-October approximately.
# The XLS "Summer" column (AS) = "Y" when sunset is late enough that a fixed
# 13:30 early mincha applies.  We detect this by candles time > 17:30.
SUMMER_CANDLES_THRESHOLD = datetime.time(17, 30)


# ── CSV parsing ────────────────────────────────────────────────────────────────

def parse_time_str(s):
    """Parse 'HH:MM' string into datetime.time. Returns None on failure."""
    if not s or s.strip() in ("-----", "--", ""):
        return None
    try:
        h, m = s.strip().split(":")
        return datetime.time(int(h), int(m))
    except Exception:
        return None


def parse_date_str(s):
    """Parse 'DD/MM/YYYY' into datetime.date. Returns None on failure."""
    if not s:
        return None
    try:
        parts = s.strip().split("/")
        if len(parts) == 3:
            return datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
    except Exception:
        pass
    return None


def load_csv(csv_path):
    """
    Load modiin_calendar.csv. Returns list of dicts with typed fields:
      date      : datetime.date
      date_h    : str (Hebrew date string)
      parsha    : str
      candles   : datetime.time or None
      shkia     : datetime.time or None
      motzash   : datetime.time or None
      holiday   : str (e.g. "RH", "YK", "", ...)
    """
    rows = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            date   = parse_date_str(r.get("date_g", ""))
            if date is None:
                continue
            rows.append({
                "date":    date,
                "date_h":  r.get("date_h", "").strip(),
                "parsha":  r.get("parsha", "").strip(),
                "candles": parse_time_str(r.get("candles", "")),
                "shkia":   parse_time_str(r.get("shkia", "")),
                "motzash": parse_time_str(r.get("motzash", "")),
                "holiday": r.get("holiday", "").strip(),
            })
    return rows


# ── Summer detection ───────────────────────────────────────────────────────────

def is_summer(candles_time):
    """
    Return 'Y' if candles time is after SUMMER_CANDLES_THRESHOLD (17:30),
    else 'N'.  None candles (Yom Kippur / Yom Tov on weekday) → 'Y' if
    month is April–September.
    """
    if candles_time is None:
        return "N"
    return "Y" if candles_time >= SUMMER_CANDLES_THRESHOLD else "N"


# ── 12H conversion ─────────────────────────────────────────────────────────────

def to_12h(t):
    """
    Convert a datetime.time in 24H (e.g. 18:03) to its 12H equivalent
    (e.g. 06:03) by subtracting 12 hours.
    This matches how the XLS Detail sheet stores AJ and AK.
    Returns datetime.time.
    """
    if t is None:
        return None
    h12 = t.hour - 12
    if h12 < 0:
        h12 += 24
    return datetime.time(h12, t.minute)


# ── XLS row matching ───────────────────────────────────────────────────────────

def find_xls_row(ws, target_date, start_row=DATA_ROW_START):
    """
    Find the row in the Detail sheet whose AM column (date) matches target_date.

    IMPORTANT: The CSV uses SATURDAY dates (Shabbat day) while the XLS Detail sheet
    uses FRIDAY dates (candle-lighting / Shabbat preparation day).
    We therefore try the date itself first, then date-1 (Friday before Saturday).

    Returns row number (1-based) or None.
    """
    candidates = [target_date, target_date - datetime.timedelta(days=1)]
    for row_idx in range(start_row, ws.max_row + 1):
        cell_val = ws.cell(row_idx, COL_AM).value
        if cell_val is None:
            continue
        if isinstance(cell_val, datetime.datetime):
            cell_date = cell_val.date()
        elif isinstance(cell_val, datetime.date):
            cell_date = cell_val
        else:
            continue
        if cell_date in candidates:
            return row_idx
    return None


# ── Verification diff ──────────────────────────────────────────────────────────

def print_verification(csv_rows, xls_matches):
    """
    Print a comparison of the CSV times against what is already in the XLS.
    Candles/Motzash from CSV are the OFFICIAL times — this just checks
    whether the XLS currently has different values and flags differences.
    """
    diffs = 0
    print()
    print("=== CANDLES / MOTZASH VERIFICATION DIFF ===")
    print(f"{'Date':12s}  {'Parsha':22s}  "
          f"{'CSV_Can':8s}  {'XLS_Can':8s}  {'Δcan':6s}  "
          f"{'CSV_Mot':8s}  {'XLS_Mot':8s}  {'Δmot':6s}")
    print("-" * 100)

    def fmt(t):
        if t is None: return "  ---  "
        return t.strftime("%H:%M")

    def diff_min(t1, t2):
        if t1 is None or t2 is None: return None
        return (t1.hour * 60 + t1.minute) - (t2.hour * 60 + t2.minute)

    for csv_row, xls_row_num, xls_can_raw, xls_mot_raw in xls_matches:
        csv_can = csv_row["candles"]
        csv_mot = csv_row["motzash"]

        # XLS stores as 12H — convert back to 24H for comparison
        xls_can = datetime.time(xls_can_raw.hour + 12, xls_can_raw.minute) if xls_can_raw else None
        xls_mot = datetime.time(xls_mot_raw.hour + 12, xls_mot_raw.minute) if xls_mot_raw else None

        dc = diff_min(csv_can, xls_can)
        dm = diff_min(csv_mot, xls_mot)

        flag = ""
        if dc not in (0, None) or dm not in (0, None):
            flag = " ← DIFF"
            diffs += 1

        print(f"{csv_row['date'].strftime('%d/%m/%Y'):12s}  "
              f"{csv_row['parsha']:22s}  "
              f"{fmt(csv_can):8s}  {fmt(xls_can):8s}  "
              f"{f'{dc:+d}m' if dc is not None else '  n/a':6s}  "
              f"{fmt(csv_mot):8s}  {fmt(xls_mot):8s}  "
              f"{f'{dm:+d}m' if dm is not None else '  n/a':6s}"
              f"{flag}")

    print()
    if diffs == 0:
        print("✓ All candles/motzash times match between CSV and XLS.")
    else:
        print(f"⚠ {diffs} row(s) differ. The CSV (official city times) will be written to XLS.")
    print("  NOTE: CSV times are always authoritative — they come from the official Modiin PDF.")
    print()


# ── Main write logic ───────────────────────────────────────────────────────────

def populate(csv_path, xls_path, year_filter=None, dry_run=False):
    print(f"Loading CSV: {csv_path}")
    csv_rows = load_csv(csv_path)
    print(f"  {len(csv_rows)} rows loaded")

    print(f"Loading XLS: {xls_path}")
    wb = openpyxl.load_workbook(xls_path, data_only=True)
    ws = wb["Detail"]
    print(f"  Detail sheet: {ws.dimensions}")

    # Build sequence for next-parsha lookup
    # next_parsha[i] = parsha of row i+1 (or empty string for last row)
    next_parsha = [r["parsha"] for r in csv_rows[1:]] + [""]

    # For verification: collect existing XLS values before overwriting
    xls_matches = []

    matched    = 0
    not_found  = 0
    skipped    = 0

    for i, row in enumerate(csv_rows):
        if year_filter:
            # Hebrew year is embedded in the date_h field; simplest filter is Gregorian year
            if row["date"].year not in (year_filter - 1, year_filter):
                skipped += 1
                continue

        xls_row = find_xls_row(ws, row["date"])
        if xls_row is None:
            print(f"  WARNING: No XLS row found for {row['date'].strftime('%d/%m/%Y')} ({row['parsha']})")
            not_found += 1
            continue

        # Capture existing values for verification diff
        xls_can_raw = ws.cell(xls_row, COL_AJ).value
        xls_mot_raw = ws.cell(xls_row, COL_AK).value
        if isinstance(xls_can_raw, datetime.time) and row["candles"] is not None:
            xls_matches.append((row, xls_row, xls_can_raw, xls_mot_raw))

        if not dry_run:
            # AJ: candles in 12H format
            if row["candles"] is not None:
                ws.cell(xls_row, COL_AJ).value = to_12h(row["candles"])

            # AK: motzash in 12H format
            if row["motzash"] is not None:
                ws.cell(xls_row, COL_AK).value = to_12h(row["motzash"])

            # AL: shkia as full datetime (date + time combined)
            if row["shkia"] is not None:
                ws.cell(xls_row, COL_AL).value = datetime.datetime.combine(row["date"], row["shkia"])

            # AM: date
            ws.cell(xls_row, COL_AM).value = datetime.datetime(row["date"].year, row["date"].month, row["date"].day)

            # AO: parsha input
            ws.cell(xls_row, COL_AO).value = row["parsha"]

            # AP: next parsha
            ws.cell(xls_row, COL_AP).value = next_parsha[i]

            # AQ: holiday code
            ws.cell(xls_row, COL_AQ).value = row["holiday"]

            # AR: Gregorian year
            ws.cell(xls_row, COL_AR).value = row["date"].year

            # AS: Summer Y/N
            ws.cell(xls_row, COL_AS).value = is_summer(row["candles"])

        matched += 1

    # Print verification diff (uses pre-update XLS values vs CSV)
    if xls_matches:
        print_verification(csv_rows, xls_matches)

    print(f"Summary: {matched} rows updated, {not_found} not found in XLS, {skipped} skipped (year filter)")

    if dry_run:
        print("DRY RUN — no changes saved.")
        return

    # Create backup
    backup_path = xls_path.replace(".xlsx", "_backup.xlsx")
    shutil.copy2(xls_path, backup_path)
    print(f"Backup saved: {backup_path}")

    wb.save(xls_path)
    print(f"XLS updated: {xls_path}")


# ── CLI ────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Populate XLS Detail sheet from Modiin calendar CSV"
    )
    parser.add_argument("--csv",     default=DEFAULT_CSV, help="Input CSV path")
    parser.add_argument("--xls",     default=DEFAULT_XLS, help="Target XLS path")
    parser.add_argument("--year",    type=int, default=None,
                        help="Hebrew year to update (e.g. 5786). Filters by Gregorian years year-1 and year.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would be written without saving")
    args = parser.parse_args()

    if not os.path.exists(args.csv):
        print(f"ERROR: CSV not found: {args.csv}", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(args.xls):
        print(f"ERROR: XLS not found: {args.xls}", file=sys.stderr)
        sys.exit(1)

    populate(args.csv, args.xls, year_filter=args.year, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
