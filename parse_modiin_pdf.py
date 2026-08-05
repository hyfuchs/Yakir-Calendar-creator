"""
parse_modiin_pdf.py
===================
Parse the Modiin annual candles/shkia/motzash PDF into a structured CSV.
This CSV is the input for populate_xls.py.

Usage:
    python parse_modiin_pdf.py                          # uses default PDF path
    python parse_modiin_pdf.py --pdf path/to/file.pdf  # explicit PDF
    python parse_modiin_pdf.py --out output.csv         # explicit CSV output

Output CSV columns:
    date_g    - Gregorian date DD/MM/YYYY
    date_h    - Hebrew date (e.g. "כז אלול תשפ"ה")
    parsha    - Parasha / holiday name. Ordinary Shabbatot use the PDF text
                (double-parshiot normalized, see below). On a yom-tov that displaces
                the weekly reading (ראש השנה, יום כיפור, סוכות, שמיני עצרת, פסח,
                שבועות) the PDF text is unreliable, so the name is derived from the
                date via pyluach (Israel schedule) instead. See holiday_name_for_date().
    candles   - הדלקת נרות (HH:MM or "-----")
    shkia     - שקיעה sunset (HH:MM)
    motzash   - מוצאי שבת (HH:MM or "-----")
    is_shabbat  - True if this is a regular Shabbat row
    holiday   - holiday code if detected (RH, YK, SUK, SIM, SMI, PES, SHV, empty)

Double-parsha note: the PDF extracts combined readings in reversed RTL order with
messy whitespace (e.g. 'וילך   -נצבים'). normalize_parsha() reverses them to correct
Torah reading order with a clean hyphen ('נצבים-וילך'). Runs automatically — next
year's PDF needs no manual fix. See RULES.md, Tool 1 section.
"""

import argparse
import csv
import datetime
import os
import re
import sys

try:
    import fitz  # pymupdf
except ImportError:
    print("ERROR: pymupdf not installed. Run: pip install pymupdf", file=sys.stderr)
    sys.exit(1)

# pyluach is OPTIONAL here: it's used only to give holiday Shabbatot a correct
# parsha name (the PDF text is unreliable on those, e.g. 'א' for ראש השנה). If it's
# missing we WARN and fall back to the raw PDF text — the parser must never
# hard-fail on this enhancement (unlike make_skeleton_csv.py, which requires it).
try:
    from pyluach import dates as _hdates, parshios as _parshios
    _HAVE_PYLUACH = True
except ImportError:
    _HAVE_PYLUACH = False
    print("WARNING: pyluach not installed; holiday-Shabbat parsha names fall back "
          "to raw PDF text. Run: pip install pyluach", file=sys.stderr)

# ── default paths ──────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PDF  = os.path.join(SCRIPT_DIR,
    "Yakir-calendar-2025-6", "Modiin-2025-6-Candels.pdf")
DEFAULT_CSV  = os.path.join(SCRIPT_DIR, "modiin_calendar.csv")

# ── column x-ranges (from PDF layout analysis) ────────────────────────────────
# PDF is RTL; columns from right (high x) to left (low x):
#   date_h  x > 375
#   date_g  x 327–374
#   parsha  x 230–326
#   candles x 185–229  (or "-----")
#   shkia   x 140–184
#   motzash x 70–139

COL_MOTZASH_MAX = 139
COL_SHKIA_MAX   = 184
COL_CANDLES_MAX = 229
COL_PARSHA_MAX  = 326
COL_DATEG_MAX   = 374
# date_h: x > 374

DATA_Y_MIN = 115   # skip title/header rows

# ── holiday detection ──────────────────────────────────────────────────────────
HOLIDAY_MAP = {
    "ראש השנה":     "RH",
    "דראש השנה":    "RH",
    "יום כיפור":    "YK",
    "סוכות":        "SUK",
    "שמחת תורה":    "SIM",
    "שמיני עצרת":   "SMI",
    "פסח":          "PES",
    "שביעי של פסח": "PES",
    "שבועות":       "SHV",
}

def normalize_parsha(parsha):
    """
    Fix double-parsha strings extracted from the PDF.

    Hebrew is RTL, so a combined reading like 'נצבים-וילך' (Nitzavim read first,
    then Vayelech) is stored/extracted in reversed logical order with irregular
    whitespace, e.g. 'וילך   -נצבים'. Split on the dash, strip, reverse to Torah
    reading order, and rejoin with a single clean hyphen.

    Single parshiot (no dash) are returned unchanged (aside from trimming).
    """
    s = (parsha or "").strip()
    if "-" not in s:
        return s
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) != 2:
        return s  # unexpected shape — leave as-is
    return f"{parts[1]}-{parts[0]}"


def detect_holiday(parsha):
    if not parsha:
        return ""
    for keyword, code in HOLIDAY_MAP.items():
        if keyword in parsha:
            return code
    return ""


def holiday_name_for_date(date_g):
    """Return the pyluach (Israel-schedule) holiday name for a date that DISPLACES
    the weekly Torah reading, else "". Used to override the unreliable PDF parsha
    text on yom-tov Shabbatot (the PDF gave 'א' for 1 Tishrei / ראש השנה).

    SAFE PREDICATE — override only when BOTH hold:
      • the date has NO weekly parsha (getparsha_string is None), AND
      • pyluach reports a holiday on it.
    This intersection isolates the displacing yom-tovim (ראש השנה, יום כיפור,
    סוכות, שמיני עצרת, פסח, שבועות) while leaving alone:
      • ordinary parshiot — different PDF spellings like שלח לך / פנחס are kept, and
      • minor holidays that fall ON a normal parsha Shabbat (Chanukah, ט״ו בשבט),
        where getparsha_string is truthy so we keep the PDF parsha.
    Returns the RAW pyluach name (no HOLIDAY_NAME_MAP remap — that belongs to the
    render layer). Robust to a missing pyluach or a malformed date → "".
    """
    if not _HAVE_PYLUACH:
        return ""
    try:
        d = datetime.datetime.strptime(date_g, "%d/%m/%Y").date()
    except (ValueError, TypeError):
        return ""
    g = _hdates.GregorianDate(d.year, d.month, d.day)
    if _parshios.getparsha_string(g, hebrew=True, israel=True) is not None:
        return ""   # has a weekly parsha (ordinary Shabbat, or Chanukah/Tu-BiShvat) → keep PDF
    hol = g.to_heb().holiday(hebrew=True, israel=True)
    return hol or ""


def is_time(s):
    """Return True if string looks like HH:MM."""
    if not s or s == "-----":
        return s == "-----" or False
    parts = s.split(":")
    return len(parts) == 2 and all(p.isdigit() for p in parts)


# ── content-based extraction ────────────────────────────────────────────────
RE_GREG   = re.compile(r"\d{2}/\d{2}/\d{4}")          # Gregorian date DD/MM/YYYY
RE_TIME   = re.compile(r"\d{1,2}:\d{2}")              # time HH:MM
RE_DASHES = re.compile(r"-{3,}")                       # "-----" (no candles/motzash)
STOP_TOKENS = ("צום", "צומות", "מקוואות", "כניסת הצום", "צאת הצום")


def _clean_quotes(s):
    """Normalize curly Hebrew quotes to straight ones (PDF-year dependent)."""
    return (s or "").replace("”", '"').replace("’", "'").strip()


def _to_min(t):
    """'HH:MM' → minutes since midnight (for chronological ordering)."""
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def parse_pdf(pdf_path, debug=False):
    """
    Content-based parser: identifies data rows by a DD/MM/YYYY date and reads
    fields by their SHAPE (date/time regex) and CHRONOLOGY, not screen x-position.
    This survives the municipality's yearly layout tweaks (verified on both the
    2025-6 layout — separate spans — and the 2026-7 layout — merged blob).

    Field logic per row:
      • Gregorian date  = the DD/MM/YYYY token.
      • The three times  = ALL HH:MM tokens in the row (standalone or embedded in a
        merged span). Assigned by CHRONOLOGY: earliest = candles (pre-sunset),
        middle = shkia (sunset), latest = motzash (after). A "-----" means that
        time is absent (e.g. no motzash / no candles on certain yom-tov rows).
      • parsha + Hebrew date = the remaining text, with the date/time tokens removed.
        parsha is the text nearest the date; the Hebrew date is the rest.

    Returns list of row dicts: date_g, date_h, parsha, candles, shkia, motzash, holiday.
    Stops at the fasts/mikvaot section.
    """
    doc  = fitz.open(pdf_path)
    page = doc[0]

    raw = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for line in b["lines"]:
            for s in line["spans"]:
                txt = s["text"].strip()
                if txt:
                    raw.append((round(s["bbox"][0]), round(s["bbox"][1]), txt))
    raw.sort(key=lambda s: (s[1], s[0]))

    # Group spans into rows by y (within 4 px).
    rows, cur_y, cur = [], None, []
    for x, y, t in raw:
        if cur_y is None or abs(y - cur_y) > 4:
            if cur:
                rows.append(cur)
            cur, cur_y = [], y
        cur.append((x, t))
    if cur:
        rows.append(cur)

    results = []
    for row in rows:
        spans  = sorted(row)                          # by x (RTL reading order)
        joined = "".join(t for _, t in spans)

        if any(tok in joined for tok in STOP_TOKENS):
            break
        gm = RE_GREG.search(joined)
        if not gm:
            continue
        date_g = gm.group(0)

        # Collect ALL times across the row's spans; assign by chronology.
        has_dashes = bool(RE_DASHES.search(joined))
        times = sorted((_to_min(t) for t in RE_TIME.findall(joined)))
        def hhmm(mins):
            return f"{mins // 60:02d}:{mins % 60:02d}"
        # Expected 3 times (candles<shkia<motzash). When one is "-----" there are 2;
        # decide which is missing by whether dashes appear and the gaps.
        candles = shkia = motzash = "-----"
        if len(times) >= 3:
            candles, shkia, motzash = hhmm(times[0]), hhmm(times[1]), hhmm(times[-1])
        elif len(times) == 2:
            # Two times + a dashes token. Most common: motzash missing (yom-tov eve,
            # candles+shkia present) → candles<shkia, motzash="-----". Occasionally
            # candles missing (candles="-----", shkia<motzash). Disambiguate by gap:
            # shkia→motzash is ~30 min; candles→shkia is ~20 min. If the two times are
            # >25 min apart, they're shkia+motzash (candles missing); else candles+shkia.
            a, b = times
            if b - a > 25:
                shkia, motzash = hhmm(a), hhmm(b)      # candles missing
            else:
                candles, shkia = hhmm(a), hhmm(b)      # motzash missing
        elif len(times) == 1:
            shkia = hhmm(times[0])

        # parsha + heb date: strip date & time tokens from the joined text, then
        # split around where the date was. Text right of the date (RTL: after it in
        # the string) is the Hebrew date; text left is the parsha.
        before = joined[:gm.start()]                  # RTL: appears visually to the right
        after  = joined[gm.end():]
        def strip_times(s):
            s = RE_TIME.sub("", s)
            s = RE_DASHES.sub("", s)
            return s.strip()
        # Layout A (2026-7): "<heb><date><parsha><candles>" → parsha in `after`.
        # Layout B (2025-6): "<times><parsha><date><heb>"   → parsha in `before`.
        cand_before = strip_times(before)
        cand_after  = strip_times(after)
        # The Hebrew date contains a month name or the year; the parsha does not
        # reliably — so pick by heuristic: the piece that looks like a Hebrew date
        # (contains a known month or "תשפ") is the heb date; the other is parsha.
        def looks_like_hebdate(s):
            return ("תשפ" in s) or any(mo in s for mo in
                    ("תשרי","חשון","כסלו","טבת","שבט","אדר","ניסן","אייר","סיון","תמוז","אב","אלול"))
        if looks_like_hebdate(cand_after) and not looks_like_hebdate(cand_before):
            heb_date, parsha_raw = cand_after, cand_before
        elif looks_like_hebdate(cand_before) and not looks_like_hebdate(cand_after):
            heb_date, parsha_raw = cand_before, cand_after
        else:
            # both or neither: fall back to layout A order (after=heb in 2026-7)
            heb_date = cand_after or cand_before
            parsha_raw = cand_before or cand_after
            # if both present and ambiguous, prefer: shorter side = parsha
            if cand_before and cand_after:
                parsha_raw, heb_date = (cand_before, cand_after)

        # Holiday-Shabbat override: on a yom-tov that displaces the weekly reading,
        # the PDF parsha text is unreliable (e.g. 'א' for ראש השנה), so use the
        # pyluach holiday name instead. Ordinary Shabbatot keep the PDF parsha
        # (including normalize_parsha's double-parsha RTL fix). Never run
        # normalize_parsha on the pyluach name (it's a single, correctly-ordered token).
        hol_name = holiday_name_for_date(date_g)
        if hol_name:
            parsha = hol_name
        else:
            parsha = normalize_parsha(_clean_quotes(parsha_raw))
        heb_date = _clean_quotes(heb_date)
        holiday  = detect_holiday(parsha)

        if debug:
            print(f"[debug] {date_g} | heb={heb_date!r} | parsha={parsha!r} | "
                  f"cand={candles} shkia={shkia} motz={motzash} | hol={holiday}")

        results.append({
            "date_g":  date_g,
            "date_h":  heb_date,
            "parsha":  parsha,
            "candles": candles,
            "shkia":   shkia,
            "motzash": motzash,
            "holiday": holiday,
        })

    return results


def save_csv(rows, csv_path):
    fieldnames = ["date_g", "date_h", "parsha", "candles", "shkia", "motzash", "holiday"]
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"CSV saved → {csv_path}  ({len(rows)} rows)")


def print_table(rows):
    """Pretty-print the parsed data with verification column."""
    import datetime

    def parse_time(s):
        if not s or s == "-----":
            return None
        try:
            h, m = s.split(":")
            return datetime.time(int(h), int(m))
        except Exception:
            return None

    def round5(t):
        """Round a datetime.time to nearest 5 minutes."""
        if t is None:
            return None
        total = t.hour * 60 + t.minute
        rounded = round(total / 5) * 5
        return datetime.time(rounded // 60, rounded % 60)

    def fmt(t):
        return t.strftime("%H:%M") if t else "-----"

    def diff_min(t1, t2):
        """Return t1 - t2 in minutes (both datetime.time)."""
        if t1 is None or t2 is None:
            return None
        d = (t1.hour * 60 + t1.minute) - (t2.hour * 60 + t2.minute)
        return d

    print()
    print(f"{'date_g':12s}  {'parsha':22s}  {'candles':8s}  {'shkia':8s}  {'motzash':8s}  "
          f"{'calc_can':9s}  {'can_diff':9s}  {'calc_mot':9s}  {'mot_diff':9s}  holiday")
    print("-" * 120)

    for r in rows:
        shkia_t   = parse_time(r["shkia"])
        candles_t = parse_time(r["candles"])
        motzash_t = parse_time(r["motzash"])

        # Standard rule: candles = shkia - 18 min, motzash = shkia + 42 min
        # (these are approximate — the PDF values are authoritative)
        calc_can = round5(
            datetime.time(
                *divmod((shkia_t.hour * 60 + shkia_t.minute - 18), 60)
            ) if shkia_t else None
        )
        calc_mot = round5(
            datetime.time(
                *divmod((shkia_t.hour * 60 + shkia_t.minute + 42), 60)
            ) if shkia_t else None
        )

        can_diff = diff_min(candles_t, calc_can)
        mot_diff = diff_min(motzash_t, calc_mot)

        can_ok = f"{can_diff:+d}min" if can_diff is not None else "  n/a  "
        mot_ok = f"{mot_diff:+d}min" if mot_diff is not None else "  n/a  "

        print(f"{r['date_g']:12s}  {r['parsha']:22s}  "
              f"{r['candles']:8s}  {r['shkia']:8s}  {r['motzash']:8s}  "
              f"{fmt(calc_can):9s}  {can_ok:9s}  "
              f"{fmt(calc_mot):9s}  {mot_ok:9s}  "
              f"{r['holiday']}")

    print()
    print("NOTE: calc_can = shkia−18min rounded to 5.  calc_mot = shkia+42min rounded to 5.")
    print("      Differences show how much the official times deviate from this simple rule.")
    print("      The CSV/XLS always uses the OFFICIAL times from the PDF — never the calculated ones.")


def main():
    parser = argparse.ArgumentParser(description="Parse Modiin candles PDF to CSV")
    parser.add_argument("--pdf", default=DEFAULT_PDF, help="Input PDF path")
    parser.add_argument("--out", default=None,
                        help="Output CSV path. If omitted, writes modiin_calendar.csv "
                             "next to the input PDF (NOT the code folder).")
    parser.add_argument("--no-csv", action="store_true", help="Skip CSV output, only print table")
    parser.add_argument("--debug", action="store_true", help="Print each parsed row (diagnosis)")
    args = parser.parse_args()

    if not os.path.exists(args.pdf):
        print(f"ERROR: PDF not found: {args.pdf}", file=sys.stderr)
        sys.exit(1)

    # Default output = modiin_calendar.csv in the SAME directory as the input PDF.
    out_path = args.out or os.path.join(os.path.dirname(os.path.abspath(args.pdf)),
                                        "modiin_calendar.csv")

    print(f"Parsing: {args.pdf}")
    rows = parse_pdf(args.pdf, debug=args.debug)
    print(f"Found {len(rows)} rows")

    print_table(rows)

    if not args.no_csv:
        save_csv(rows, out_path)


if __name__ == "__main__":
    main()
