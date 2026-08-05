#!/usr/bin/env python3
"""
make_skeleton_csv.py — generate a blank-times modiin_calendar.csv for a Hebrew year.

FALLBACK for years where only an image scan of the Modiin candles table is
available (no digital PDF to parse). Everything except the actual times is
computable, so this pre-fills dates / parsha / holiday from pyluach and leaves
the time columns (candles, shkia, motzash) BLANK for you to type from the scan.

Output schema matches parse_modiin_pdf.py exactly, so the rest of the pipeline
(modiin_weeks.py → yakir_calendar.py) runs unchanged:
    date_g, date_h, parsha, candles, shkia, motzash, holiday
plus helper columns candles_hint / shkia_hint / motzash_hint (ephem estimates,
so you can sanity-check your typing; ignored by the pipeline — delete or keep).

Usage:
    python3 make_skeleton_csv.py --year 5788 --out modiin_calendar.csv
Then open the CSV, type the official candles + motzash (+ optional shkia) times
from the scan, save, and continue with modiin_weeks.py.

Rows: every Saturday from the last Shabbat of the previous Elul through the last
Shabbat before the next Rosh Hashana (Saturdays only — that's what the pipeline
consumes). Handles leap years (Adar I/II) automatically.
"""
import argparse
import csv
import datetime
import os
import sys

try:
    from pyluach import dates as hdates, parshios, hebrewcal
except ImportError:
    sys.exit("pyluach not installed. Run: pip install pyluach")

# Reuse the project's helpers so behavior matches the PDF path exactly.
from parse_modiin_pdf import detect_holiday
try:
    from modiin_weeks import ephem_sunset_minutes, hhmm, DEFAULT_HORIZON
    _HAVE_EPHEM = True
except Exception:
    _HAVE_EPHEM = False

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(SCRIPT_DIR, "modiin_calendar.csv")

FIELDNAMES = ["date_g", "date_h", "parsha", "candles", "shkia", "motzash", "holiday",
              "candles_hint", "shkia_hint", "motzash_hint"]

# ephem hint offsets (rough — the official typed times always win)
CANDLES_BEFORE_SUNSET = 20    # candles ≈ sunset − 20 min
MOTZASH_AFTER_SUNSET  = 30    # motzash ≈ sunset + 30 min


def _greg(hebdate):
    g = hebdate.to_greg()
    return datetime.date(g.year, g.month, g.day)


def parsha_or_holiday(greg_date):
    """Israel-schedule parsha for a Shabbat, or the holiday name if one overrides
    it. pyluach already returns double-parshiot in correct Torah order
    ('נצבים, וילך'), so we just swap the comma for a dash — do NOT run
    normalize_parsha (that reverses order, which is only needed for the PDF path
    where the source is reversed)."""
    g = hdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day)
    p = parshios.getparsha_string(g, hebrew=True, israel=True)
    if p:
        return p.replace(", ", "-").replace(",", "-")
    hol = g.to_heb().holiday(hebrew=True, israel=True)
    return hol or ""


def heb_date_str(greg_date):
    h = hdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    return f"{h.hebrew_day()} {h.month_name(hebrew=True)} {hebrewcal.Year(h.year).year_string()}"


def _hint(sunday_sat_date, delta):
    if not _HAVE_EPHEM:
        return ""
    ss = ephem_sunset_minutes(sunday_sat_date, DEFAULT_HORIZON)
    return hhmm(ss + delta)


def build_rows(heb_year):
    """All Saturdays from the last Shabbat of the previous Elul through the last
    Shabbat before 1 Tishrei of heb_year+1."""
    start = _greg(hdates.HebrewDate(heb_year, 7, 1))       # 1 Tishrei (Rosh Hashana)
    nextrh = _greg(hdates.HebrewDate(heb_year + 1, 7, 1))  # next Rosh Hashana
    # First row = the last Shabbat BEFORE Rosh Hashana (the previous-Elul tail the
    # pipeline expects for its week-label lookahead): the Saturday within the 7 days
    # before 1 Tishrei.
    d = start - datetime.timedelta(days=1)
    while d.weekday() != 5:      # 5 = Saturday, walking back from RH-eve
        d -= datetime.timedelta(days=1)
    rows = []
    # Include Saturdays through the first Shabbat on/after next Rosh Hashana, so the
    # trailing boundary matches the PDF and the next year's lookahead has its anchor.
    end = nextrh + datetime.timedelta(days=7)
    while d < end:
        rows.append({
            "date_g":  d.strftime("%d/%m/%Y"),
            "date_h":  heb_date_str(d),
            "parsha":  parsha_or_holiday(d),
            "candles": "", "shkia": "", "motzash": "",
            "holiday": detect_holiday(parsha_or_holiday(d)),
            "candles_hint": _hint(d, -CANDLES_BEFORE_SUNSET),
            "shkia_hint":   _hint(d, 0),
            "motzash_hint": _hint(d, +MOTZASH_AFTER_SUNSET),
        })
        d += datetime.timedelta(days=7)
    return rows


def main():
    ap = argparse.ArgumentParser(description="Generate a blank-times modiin_calendar.csv for a Hebrew year")
    ap.add_argument("--year", type=int, required=True, help="Hebrew year, e.g. 5788")
    ap.add_argument("--out",  default=None,
                    help="Output CSV path. If omitted, writes modiin_calendar.csv to the "
                         "current directory (run this from your year folder).")
    args = ap.parse_args()

    # Default output = modiin_calendar.csv in the CURRENT working directory (this
    # tool has no input file to anchor to — run it from the target year folder).
    out_path = args.out or os.path.join(os.getcwd(), "modiin_calendar.csv")

    rows = build_rows(args.year)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(rows)

    print(f"Skeleton saved → {os.path.abspath(out_path)}  ({len(rows)} Saturday rows)")
    print("Next: open it, type the official candles + motzash (+ optional shkia) from the scan,")
    print("      then run modiin_weeks.py. The *_hint columns are ephem estimates to check against.")
    if not _HAVE_EPHEM:
        print("  (ephem not available — hint columns left blank)")


if __name__ == "__main__":
    main()
