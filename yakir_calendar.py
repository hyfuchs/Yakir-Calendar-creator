"""
yakir_calendar.py  –  v3
========================
Generates one landscape A4 page per Jewish month, matching the reference layout.

CRITICAL: Candles & Motzash values come DIRECTLY from the XLS. Never calculated here.

Usage:
    python yakir_calendar.py                        # all months
    python yakir_calendar.py --months 0 1 2         # specific months (0-based)
    python yakir_calendar.py --out my_output.docx
    python yakir_calendar.py --pdf                  # also try to produce PDF
"""

import argparse, csv, datetime, os, shutil, subprocess, sys

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("ERROR: pip install python-docx", file=sys.stderr); sys.exit(1)
try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl", file=sys.stderr); sys.exit(1)
try:
    from pyluach import dates as _hebdates, parshios as _parshios, hebrewcal as _hebcal
    from pyluach.parshios import four_parshios as _four_parshios
except ImportError:
    print("ERROR: pip install pyluach", file=sys.stderr); sys.exit(1)
# Reuse the calibrated ephem weekday-mincha logic from modiin_weeks.py so the
# year-boundary fallback matches the CSV pipeline exactly (same rule/constants).
try:
    from modiin_weeks import (ephem_sunset_minutes, floor5, hhmm,
                              WEEKDAY_MINCHA_OFFSET_MIN, DEFAULT_HORIZON)
    _HAVE_EPHEM = True
except Exception:
    _HAVE_EPHEM = False

# ── paths ──────────────────────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
BASE_DIR    = os.path.join(SCRIPT_DIR, "..")
DEFAULT_XLS  = os.path.join(BASE_DIR, "yakircal master v61 Added shkia+14 times starting 5785.xlsx")
# Input defaults to the CURRENT working directory (never the code folder); output
# goes next to the input --weeks-csv.
DEFAULT_WEEKS_CSV = os.path.join(os.getcwd(), "modiin_weeks.csv")
DEFAULT_LOGO = os.path.join(SCRIPT_DIR, "yakir_logo.png")

# ── page: A4 landscape ─────────────────────────────────────────────────────────
PG_W, PG_H = 29.7, 21.0
MAR   = 1.0    # left/right margin (cm)
MAR_V = 0.6    # top/bottom margin (cm) — tighter to fit one page
BODY_W = PG_W - 2 * MAR   # 27.7 cm

# ── fonts ──────────────────────────────────────────────────────────────────────
# CRITICAL: the .docx does NOT embed fonts — it only NAMES them, and Word renders
# with whatever that name resolves to on the machine that OPENS the file (always
# Windows Word here). So the font name we write must be IDENTICAL regardless of
# which machine generated the file. Do NOT branch on the build host — that was a
# real bug: WSL builds named "Arial" (correct) while native-Windows builds named
# "DejaVu Sans" (wide/absent Hebrew glyphs → the long week label wrapped → the
# month spilled to 2 pages). Arial is on Windows + Mac; on bare Linux LibreOffice
# substitutes Liberation Sans, which is metric-compatible with Arial, so line
# widths — and therefore pagination — stay the same everywhere.
F_ARIAL   = "Arial"
F_CALIBRI = "Calibri"

# ── unit helpers  (CORRECT: 1 pt = 20 twips) ──────────────────────────────────
def pt2tw(pt):  return str(int(pt * 20))
def cm2tw(cm):  return str(int(Cm(cm).pt * 20))

# ── XML primitives ─────────────────────────────────────────────────────────────

def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e

def _bidi_para(para):
    # OOXML requires pPr children in schema order: bidi must come BEFORE spacing,
    # ind and jc. Appending it last (after jc) makes Word silently ignore it → the
    # paragraph is treated LTR and right-alignment renders wrong. Insert bidi before
    # the EARLIEST-positioned of any spacing/ind/jc already present.
    pPr = para._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is not None:
        return
    bidi = _el("w:bidi", **{"w:val": "1"})
    after_tags = {qn("w:spacing"), qn("w:ind"), qn("w:jc")}
    anchor = None
    for child in pPr:                     # first child (lowest index) that bidi must precede
        if child.tag in after_tags:
            anchor = child
            break
    if anchor is not None:
        anchor.addprevious(bidi)
    else:
        pPr.insert(0, bidi)

def _spacing(para, before=0, after=0, line=None):
    kw = {"w:before": pt2tw(before), "w:after": pt2tw(after)}
    if line is not None:
        kw["w:line"] = pt2tw(line)
        kw["w:lineRule"] = "exact"
    para._p.get_or_add_pPr().append(_el("w:spacing", **kw))

def _run(para, text, font, size, bold=False, rtl=True, color=None):
    r = para.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    rPr = r._r.get_or_add_rPr()
    rf = _el("w:rFonts", **{"w:ascii": font, "w:hAnsi": font, "w:cs": font})
    rPr.insert(0, rf)
    # Complex-script size: Word sizes Hebrew (a complex script) via w:szCs, NOT
    # w:sz. python-docx's font.size only sets w:sz, so Hebrew would fall back to
    # the default 11pt. Set w:szCs explicitly so Hebrew honors the size.
    half_pt = str(int(size * 2))   # OOXML sizes are in half-points
    rPr.append(_el("w:szCs", **{"w:val": half_pt}))
    if bold:
        rPr.append(_el("w:bCs", **{"w:val": "1"}))   # bold for complex script too
    if color:
        rPr.append(_el("w:color", **{"w:val": color}))
    if rtl:
        rPr.append(_el("w:rtl", **{"w:val": "1"}))
    return r

def _cell_write(cell, text, font, size, bold=False,
                align=WD_ALIGN_PARAGRAPH.CENTER,
                before=1, after=1, line=None, rtl=True, valign="center"):
    """Clear cell and write one paragraph.
    align=None → don't emit w:jc at all; for an RTL (bidi) paragraph this yields
    natural right alignment, which Word renders correctly (an explicit jc=right on
    a bidi paragraph renders LEFT — confirmed against a hand-fixed reference file)."""
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    _spacing(p, before, after, line)
    if rtl:
        _bidi_para(p)
    _run(p, text, font, size, bold=bold, rtl=rtl)
    # vertical align
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.append(_el("w:vAlign", **{"w:val": valign}))

def _cell_para(cell, text, font, size, bold=False,
               align=WD_ALIGN_PARAGRAPH.CENTER,
               before=1, after=1, rtl=True):
    """Add a paragraph to a cell (doesn't clear existing content)."""
    p = cell.add_paragraph()
    p.alignment = align
    _spacing(p, before, after)
    if rtl:
        _bidi_para(p)
    _run(p, text, font, size, bold=bold, rtl=rtl)
    return p

def _cell_bg(cell, hex6):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(_el("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": hex6}))

def _cell_w(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(_el("w:tcW", **{"w:w": cm2tw(cm), "w:type": "dxa"}))

def _row_h(row, cm, exact=True):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el("w:trHeight",
                    **{"w:val": cm2tw(cm), "w:hRule": "exact" if exact else "atLeast"}))

def _tbl_w(tbl_el, cm):
    tblPr = tbl_el.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = _el("w:tblPr"); tbl_el.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = _el("w:tblW"); tblPr.append(tblW)
    tblW.set(qn("w:w"),    cm2tw(cm))
    tblW.set(qn("w:type"), "dxa")

def _tbl_no_borders(table):
    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = _el("w:tblPr"); table._tbl.insert(0, tblPr)
    tb = _el("w:tblBorders")
    for s in ("top","left","bottom","right","insideH","insideV"):
        tb.append(_el(f"w:{s}", **{"w:val": "none"}))
    tblPr.append(tb)

def _tbl_bidi(table):
    """Make table flow RTL visually: col 0 = rightmost column."""
    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = _el("w:tblPr"); table._tbl.insert(0, tblPr)
    tblPr.append(_el("w:bidiVisual", **{"w:val": "1"}))

def _tbl_borders(table, color="888888", sz=4):
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcB  = _el("w:tcBorders")
            for s in ("top","left","bottom","right"):
                tcB.append(_el(f"w:{s}",
                               **{"w:val": "single", "w:sz": str(sz),
                                  "w:space": "0", "w:color": color}))
            tcPr.append(tcB)

def _zero_para(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    _spacing(p, 0, 0, 1)
    return p

def _page_break(doc):
    p = _zero_para(doc)
    r = p.add_run()
    r._r.append(_el("w:br", **{"w:type": "page"}))

# ── time formatter ─────────────────────────────────────────────────────────────

def ft(v, empty=""):
    if v is None: return empty
    if isinstance(v, datetime.time):     return v.strftime("%H:%M")
    if isinstance(v, datetime.datetime): return v.strftime("%H:%M")
    s = str(v).strip()
    return empty if s in ("--","---","None","") else s

# ── weekday-mincha CSV loader ────────────────────────────────────────────────

def load_weekday_mincha(weeks_csv_path):
    """
    Load weekday mincha times from modiin_weeks.csv, keyed by the week's Sunday
    ('from' column, DD/MM/YYYY). Value = weekday_mincha_adj if set, else _calc.

    This is the CSV source that is replacing the XLS-computed weekday mincha.
    See RULES.md §2. Returns {} if the file is missing (caller falls back to XLS).

    Join alignment: the calendar shows, on Shabbat row D, the mincha for the
    week that STARTS the next day. So the caller looks up key = (D + 1 day),
    which is the Sunday of that week and equals this CSV row's 'from' value.
    """
    if not os.path.exists(weeks_csv_path):
        return {}
    table = {}
    with open(weeks_csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            key = (row.get("from") or "").strip()
            val = (row.get("weekday_mincha_adj") or "").strip() \
                or (row.get("weekday_mincha_calc") or "").strip()
            if key and val:
                table[key] = val
    return table


# ── Shabbat-block constants (see RULES.md §5) ────────────────────────────────
SHABBAT_MINCHA_GEDOLA = "13:30"   # always shown
SHARBAT_MINCHA        = "17:15"   # summer early mincha, shown when sharbat_mincha == "yes"
SHACHARIT_1_DEFAULT   = "6:45"    # shacharit1 column overrides when filled
EREV_SHABBAT_GAP_MIN  = 10        # mincha erev shabbat = resolved mincha ktana + 10 min

# ── Weekday-table constants (confirmed constant across the year in the XLS) ──
WEEKDAY_DAF    = "19:45"   # daf yomi time
WEEKDAY_ARVIT2 = "20:30"   # weekday arvit 2 (later minyan)

# Weekday shacharit constants (from reference PDF YAKIR-5786-11-v1).
# Regular week + Rosh Chodesh variants.
SHACH_WK_YOM_AH     = "06:15"        # Sun–Thu regular
SHACH_WK_FRI        = "8:30 / 6:45"  # Friday regular (2 minyanim)
SHACH_WK_RC_YOM_AH  = "06:05"        # Sun–Thu on Rosh Chodesh
SHACH_WK_RC_FRI     = "8:30 / 06:30" # Friday on Rosh Chodesh


def _plus_minutes(hhmm_str, delta):
    """Add `delta` minutes to an 'HH:MM' string. Returns '' if input is blank/invalid."""
    s = (hhmm_str or "").strip()
    if not s or ":" not in s:
        return ""
    try:
        h, m = s.split(":")
        total = (int(h) * 60 + int(m) + delta) % (24 * 60)
        return f"{total // 60:02d}:{total % 60:02d}"
    except ValueError:
        return ""


def load_shabbat_block(weeks_csv_path):
    """
    Load per-Shabbat prayer times from modiin_weeks.csv, keyed by weekend_date
    (DD/MM/YYYY). Each value is a dict with resolved (adj|calc) times plus the
    summer flag, ready to render. Returns {} if the file is missing (caller
    falls back to XLS). See RULES.md §5.

    mincha_erev (Friday) is NOT a stored column — it is derived here as the
    resolved mincha ktana + 10 min, so a ktana override shifts erev automatically
    (matches the old calendar: erev = ktana + 10, verified 50/50 weeks).
    """
    if not os.path.exists(weeks_csv_path):
        return {}
    def pick(row, base):   # adj wins over calc
        return (row.get(base + "_adj") or "").strip() or (row.get(base + "_calc") or "").strip()
    table = {}
    with open(weeks_csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            key = (row.get("weekend_date") or "").strip()
            if not key:
                continue
            ktana = pick(row, "shabbat_mincha_ktana")
            table[key] = {
                "taarich":       (row.get("taarich") or "").strip(),
                "candles":       (row.get("candles") or "").strip(),
                "motzash":       (row.get("motzash") or "").strip(),
                "mincha_ktana":  ktana,
                "mincha_erev":   _plus_minutes(ktana, EREV_SHABBAT_GAP_MIN),  # derived
                "shma":          pick(row, "sof_zman_shma"),
                "shacharit1":    (row.get("shacharit1") or "").strip() or SHACHARIT_1_DEFAULT,
                "sharbat":       (row.get("sharbat_mincha") or "").strip().lower() == "yes",
            }
    return table


def format_shabbat_mincha(block, suppress_sharbat=False):
    """Build the מנחה שבת display string from a Shabbat-block dict:
    gedola 13:30, summer Sharbat 17:15 (if applicable), then mincha ktana.
    Mirrors the XLS min2 format ('13:30  17:15  HH:MM' / '13:30    HH:MM').
    suppress_sharbat=True drops the 17:15 (festival Shabbatot — Sukkot etc. —
    have no Sharbat summer mincha)."""
    parts = [SHABBAT_MINCHA_GEDOLA]
    if block.get("sharbat") and not suppress_sharbat:
        parts.append(SHARBAT_MINCHA)
    if block.get("mincha_ktana"):
        parts.append(block["mincha_ktana"])
    return "  ".join(parts)


# ── CSV month loader (primary source, replaces XLS) ──────────────────────────

def _clean_parsha(name):
    """Strip bookkeeping suffixes ('- מברכים' etc.) from a parsha name for use
    in the week label, mirroring the XLS wkLbl format."""
    s = (name or "").strip()
    for suf in ("מברכים",):
        if s.endswith(suf):
            s = s[: -len(suf)].rstrip(" -־").strip()
    return s


def next_shabbat_label(greg_date):
    """
    Given a Shabbat date, return the display name of the *following* Shabbat,
    used to complete a week label at the dataset's year boundary (no next CSV row).

    Uses the ISRAEL parsha schedule (pyluach israel=True) — the diaspora schedule
    diverges from the Modiin calendar by up to a week. If a holiday falls on that
    Shabbat, pyluach returns no parsha, so we use the holiday name (e.g. ראש השנה).
    Double-parshiot come back comma-separated ('נצבים, וילך') → normalize to '-'.
    """
    nxt = greg_date + datetime.timedelta(days=7)
    g = _hebdates.GregorianDate(nxt.year, nxt.month, nxt.day)
    p = _parshios.getparsha_string(g, hebrew=True, israel=True)
    if p:
        return p.replace(", ", "-").replace(",", "-")
    hol = g.to_heb().holiday(hebrew=True)
    return hol or ""


def weekday_mincha_for_week(sunday_date):
    """
    Compute the weekday mincha for the week starting on sunday_date, using the
    SAME calibrated ephem rule as modiin_weeks.py: floor_to_5(earliest Sun–Thu
    sunset − 12 min). Used as a fallback when the CSV has no row for that week
    (e.g. the year boundary). Returns '' if the ephem helpers are unavailable.
    """
    if not _HAVE_EPHEM:
        return ""
    sunsets = [ephem_sunset_minutes(sunday_date + datetime.timedelta(days=i), DEFAULT_HORIZON)
               for i in range(5)]   # Sun..Thu
    earliest = int(round(min(sunsets)))
    return hhmm(floor5(earliest - WEEKDAY_MINCHA_OFFSET_MIN))


# ── month events (holidays + Rosh Chodesh) — see RULES.md §Events ────────────
# Hebrew weekday names, Sunday-based (Python weekday(): Mon=0..Sun=6 → (wd+1)%7).
HEB_DOW = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"]

# pyluach holiday names remapped to the congregation's preferred wording.
HOLIDAY_NAME_MAP = {"שמיני עצרת": "שמחת תורה"}   # Israel: same day as Simchat Torah

# Major holidays shown in the events section (this phase). Fasts are deferred:
# kept in EVENT_FASTS (unused) so they can be enabled with a one-line change.
EVENT_HOLIDAYS = {"ראש השנה", "יום כיפור", "סוכות", "שמיני עצרת",
                  "פסח", "שבועות", "חנוכה", "פורים", "שושן פורים",
                  "ל״ג בעומר", "ט״ו בשבט", "ט״ו באב", "פסח שני"}
EVENT_FASTS    = {"צום גדליה", "י׳ בטבת", "תענית אסתר", "י״ז בתמוז", "ט׳ באב"}


def _heb_dow(greg_date):
    """Hebrew weekday name for a date (Sunday-based)."""
    return HEB_DOW[(greg_date.weekday() + 1) % 7]


def _greg(hebdate):
    """pyluach HebrewDate → datetime.date."""
    g = hebdate.to_greg()
    return datetime.date(g.year, g.month, g.day)


# Shabbat Chol HaMoed: the PDF labels these with the Torah-reading name "ויאמר משה",
# which is confusing in the parsha column. Detect via pyluach and relabel + set the
# special shacharit-1 time. Short form חוה"מ fits the narrow parsha column.
def chol_hamoed_shabbat(greg_date):
    """If greg_date is a Shabbat Chol HaMoed of Sukkot/Pesach, return
    {'name': ..., 'shach1': ...}; else None. Uses pyluach israel schedule.
    Sukkot Chol HaMoed = Tishrei 16-21; Pesach Chol HaMoed = Nissan 16-20
    (excludes yom-tov first/last days)."""
    if greg_date.weekday() != 5:   # 5 = Saturday
        return None
    h = _hebdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    hol = h.holiday(hebrew=True, israel=True)
    if hol == "סוכות" and 16 <= h.day <= 21:
        return {"name": 'שבת חוה"מ סוכות', "shach1": "6:15"}
    if hol == "פסח" and 16 <= h.day <= 20:
        return {"name": 'שבת חוה"מ פסח', "shach1": "6:30"}
    return None


def special_shabbat_tag(greg_date):
    """Return the four-parshiot tag (שקלים/זכור/פרה/החודש) for a Shabbat, else ''.
    Only the four parshiot pyluach detects for free — no manual date logic."""
    h = _hebdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    return _four_parshios(h, hebrew=True) or ""


def sukkot_shabbat(greg_date):
    """True if this Shabbat falls during Sukkot (1st day OR chol hamoed),
    Tishrei 15-21. shacharit-1 is earlier (6:15) because Kohelet is read and it's
    long; and there is NO Sharbat summer 17:15 mincha on these Shabbatot.
    (pyluach month: Tishrei = 7.)"""
    if greg_date.weekday() != 5:   # 5 = Saturday
        return False
    h = _hebdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    return h.month == 7 and 15 <= h.day <= 21


def yomtov_no_minyan_shabbat(greg_date):
    """True for a festival day on Shabbat whose regular minyan schedule does NOT
    apply, so shacharit (both) and Shabbat mincha are left BLANK for manual fill:
    Shemini Atzeret / Simchat Torah (Tishrei 22) and Pesach 1st day (Nissan 15).
    (pyluach months: Tishrei = 7, Nissan = 1.)"""
    if greg_date.weekday() != 5:
        return False
    h = _hebdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    return (h.month == 7 and h.day == 22) or (h.month == 1 and h.day == 15)


# Conditional note texts (see RULES.md).
CHANUKAH_NOTE = "בחנוכה תפילת ערבית ראשונה כ-20 דקות אחרי השקיעה"
SCHOOL_NOTE   = "בחופשות בתי הספר, מתקיים לרוב מניין שחרית נוסף בימים א-ה בשעה 8:45"


def _pyluach_month(heb_year, tishrei_month_num):
    """Map our Tishrei-based month number (Tishrei=1) to pyluach's month number
    (Nissan=1, Tishrei=7), handling leap-year Adar I/II via the civil-order list."""
    civil = [m.month for m in _hebcal.Year(heb_year).itermonths()]  # Tishrei first
    return civil[tishrei_month_num - 1]


def _rosh_chodesh_dates(heb_year, py_mo):
    """Rosh Chodesh dates for the month (py_mo): day 1, plus day 30 of the
    previous month when that month has 30 days (→ a 2-day Rosh Chodesh).
    Returns a sorted list of 1 or 2 datetime.date."""
    month = _hebcal.Month(heb_year, py_mo)
    prev = month - 1
    ds = []
    if len(prev) == 30:
        ds.append(_greg(_hebdates.HebrewDate(prev.year, prev.month, 30)))
    ds.append(_greg(_hebdates.HebrewDate(heb_year, py_mo, 1)))
    ds.sort()
    return ds


def _rc_line(name, dates_):
    """Format a Rosh Chodesh (or Rosh Hashana) line from 1-2 dates."""
    if len(dates_) == 2:
        d1, d2 = dates_
        days = f"{d1.day}-{d2.day}.{d1.month}" if d1.month == d2.month \
               else f"{d1.day}.{d1.month}-{d2.day}.{d2.month}"
        return d1, f"{name} - {_heb_dow(d1)}-{_heb_dow(d2)} {days}"
    d1 = dates_[0]
    return d1, f"{name} - {_heb_dow(d1)} {d1.day}.{d1.month}"


def month_events(heb_year, tishrei_month_num):
    """
    Return the month's event lines (holidays + fasts + Rosh Chodesh) as formatted
    strings 'name - dow D.M', ordered by date. Uses pyluach (israel schedule). See RULES.md.

    Rosh Chodesh rules:
      - Current month's RC: shown UNLESS one of its days falls on Shabbat
        (then it's already visible in the Shabbat table row).
      - Next month's RC: always shown (same RC also appears as "current" on the
        next month's page — that duplication is intended).
      - Tishrei exception: the next month's RC that equals 1 Tishrei (Rosh Hashana)
        is shown as "ראש השנה" (this only happens on the Elul page); and it is
        omitted entirely if 1 Tishrei falls on Shabbat.
    """
    py_mo = _pyluach_month(heb_year, tishrei_month_num)
    month = _hebcal.Month(heb_year, py_mo)
    events = []   # (sort_date, text)

    # Holidays + fasts: first day only, deduped by (mapped) name. Exception: ראש השנה
    # is a 2-day yom tov → shown as a range. Fasts are single-day.
    seen = set()
    show = EVENT_HOLIDAYS | EVENT_FASTS
    for d in range(1, len(month) + 1):
        h = _hebdates.HebrewDate(heb_year, py_mo, d)
        hol = h.holiday(hebrew=True, israel=True)
        if not hol or hol not in show:
            continue
        name = HOLIDAY_NAME_MAP.get(hol, hol)
        if name in seen:
            continue
        seen.add(name)
        if name == "ראש השנה":
            rh = [_greg(_hebdates.HebrewDate(heb_year, py_mo, 1)),
                  _greg(_hebdates.HebrewDate(heb_year, py_mo, 2))]
            events.append(_rc_line("ראש השנה", rh))   # range formatter
        else:
            dt = _greg(h)
            events.append((dt, f"{name} - {_heb_dow(dt)} {dt.day}.{dt.month}"))

    # שביעי של פסח — a separate yom tov (candle-lighting) that pyluach lumps under
    # "פסח". In Nissan (pyluach month 1) inject it explicitly for 21 Nissan.
    if py_mo == 1:
        dt7 = _greg(_hebdates.HebrewDate(heb_year, 1, 21))
        events.append((dt7, f"שביעי של פסח - {_heb_dow(dt7)} {dt7.day}.{dt7.month}"))

    # Current month's Rosh Chodesh — skip if any of its days is Shabbat, OR if it
    # coincides with Rosh Hashana (ראש השנה gובר; e.g. 1 Tishrei). Rosh Hashana
    # is already shown as the holiday above.
    rc_cur = _rosh_chodesh_dates(heb_year, py_mo)
    rc_on_shabbat = any(d.weekday() == 5 for d in rc_cur)
    rc_is_rosh_hashana = (py_mo == 7)   # 1 Tishrei = Rosh Hashana
    if not rc_on_shabbat and not rc_is_rosh_hashana:
        events.append(_rc_line(f"ראש חודש {month.month_name(hebrew=True)}", rc_cur))

    # Next month's Rosh Chodesh (always shown), with the Tishrei→Rosh Hashana rule.
    nxt = month + 1
    rc_next = _rosh_chodesh_dates(nxt.year, nxt.month)
    if nxt.month == 7:   # pyluach month 7 = Tishrei → next-RC is Rosh Hashana
        rosh_hashana = _greg(_hebdates.HebrewDate(nxt.year, 7, 1))
        if rosh_hashana.weekday() != 5:   # skip if 1 Tishrei is Shabbat
            events.append(_rc_line("ראש השנה", [rosh_hashana]))
    else:
        events.append(_rc_line(f"ראש חודש {nxt.month_name(hebrew=True)}", rc_next))

    events.sort(key=lambda e: e[0])
    return [txt for _, txt in events]


def load_months_from_csv(weeks_csv_path):
    """
    Build the month list directly from modiin_weeks.csv — the primary data source
    (the XLS is validation-only). Returns the same month-dict shape build_page
    consumes, so rendering is unchanged. See RULES.md §2/§5.

    Months are grouped by hyrmo, in week order. Week label (wkLbl) is derived
    across the whole file as 'clean(parsha_this) - clean(parsha_next)', so the
    last week of a month reaches into the next month's first parsha.
    YomText / Rosh Chodesh have no CSV source yet → rendered empty.
    """
    if not os.path.exists(weeks_csv_path):
        return []
    with open(weeks_csv_path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    def display_name(weekend_date_str, weekend_name):
        """Parsha name for display. The PDF's parsha text is unreliable on holiday
        Shabbatot (e.g. it gave 'א' for 1 Tishrei / ראש השנה, and its wording drifts
        year to year), so DERIVE the name from the date via pyluach:
          • Chol HaMoed → 'שבת חוה"מ …' (Sukkot/Pesach).
          • Any other holiday-on-Shabbat → the holiday name (remapped, e.g.
            שמיני עצרת → שמחת תורה), NOT the PDF text.
        Ordinary Shabbatot keep the cleaned PDF parsha name."""
        try:
            d = datetime.datetime.strptime(weekend_date_str, "%d/%m/%Y").date()
            chm = chol_hamoed_shabbat(d)
            if chm:
                return chm["name"]
            h = _hebdates.GregorianDate(d.year, d.month, d.day).to_heb()
            hol = h.holiday(hebrew=True, israel=True)
            if hol:   # holiday overrides the (unreliable) PDF parsha text
                return HOLIDAY_NAME_MAP.get(hol, hol)
        except ValueError:
            pass
        return _clean_parsha(weekend_name)

    # Precompute week labels across the entire sequence (needs the next row).
    # At the dataset's end (no next row), look ahead one Shabbat via pyluach so the
    # final week still gets a range label (e.g. 'נצבים-וילך - ראש השנה').
    labels = {}
    for i, r in enumerate(rows):
        wd = (r.get("weekend_date") or "").strip()
        this_p = display_name(wd, r.get("weekend_name"))
        if i + 1 < len(rows):
            nrow = rows[i + 1]
            next_p = display_name((nrow.get("weekend_date") or "").strip(),
                                  nrow.get("weekend_name"))
        else:
            try:
                d = datetime.datetime.strptime(wd, "%d/%m/%Y").date()
                next_p = next_shabbat_label(d)
            except ValueError:
                next_p = ""
        labels[wd] = f"{this_p} - {next_p}" if next_p else this_p

    months, order, by_key = [], [], {}
    for r in rows:
        hyrmo = (r.get("hyrmo") or "").strip()
        if not hyrmo:
            continue
        wd = (r.get("weekend_date") or "").strip()
        try:
            dt = datetime.datetime.strptime(wd, "%d/%m/%Y").date()
        except ValueError:
            continue
        sb = {
            "date":    dt,
            "parsha":  display_name(wd, r.get("weekend_name")),
            "wkLbl":   labels.get(wd, ""),
            # Shabbat-block and weekday-mincha values are resolved later by
            # load_shabbat_block / load_weekday_mincha (keyed by date). Weekday
            # daf/arvit2 are constants applied at render time.
        }
        if hyrmo not in by_key:
            # Compute the events line (holidays + Rosh Chodesh) for this month.
            yomtext = ""
            try:
                hy = int(r.get("heb_year"))
                hm = int(r.get("heb_month_num"))
                yomtext = " | ".join(month_events(hy, hm))
            except (TypeError, ValueError):
                pass
            m = {
                "hyrmo":   hyrmo,
                "chodesh": (r.get("heb_month_name") or "").strip(),
                "shana":   (r.get("heb_year_str") or "").strip(),
                "shabbatot": [],
                "YomText": yomtext, "ThisRC": "", "NextRC": "",
            }
            by_key[hyrmo] = m
            order.append(hyrmo)
        by_key[hyrmo]["shabbatot"].append(sb)
    return [by_key[k] for k in order]


# ── XLS loader ─────────────────────────────────────────────────────────────────

def load_months(xls_path):
    """
    Load all months from ResultForWordMerge sheet.

    Weekday mincha/arvit times come directly from the pre-calculated wk* fields
    in the ResultForWordMerge sheet (inside the master XLS workbook).

    IMPORTANT - wkMin shift rule:
      The wkLbl "A - B" describes the week AFTER Shabbat A.
      The wkMin stored on Shabbat A's row is the mincha for the week
      STARTING after Shabbat A — i.e. it is already the correct value
      to display for row A.
      Exception: the FIRST shabbat of a month uses lastmo-wkMin (from the
      last Shabbat of the previous month).

    Source validity: ResultForWordMerge is pre-calculated for years 5785-5790
    (approx. 2024-2030). For years beyond 2030, the XLS must be extended.
    See RULES.md §5 for details.
    """
    wb  = openpyxl.load_workbook(xls_path, data_only=True)
    ws  = wb["ResultForWordMerge"]
    hdr = [c.value for c in ws[1]]
    months = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(hdr, row))
        if d.get("a-hyrmo") is None: break
        shabbatot = []
        for p in "abcde":
            if d.get(f"{p}-hyrmo") is None: break
            sb = {}
            for k in hdr:
                if k and k.startswith(f"{p}-"):
                    sb[k[2:]] = d[k]
            shabbatot.append(sb)
        lastmo = {k[7:]: d[k] for k in hdr if k and k.startswith("lastmo-")}
        months.append({
            "hyrmo":     d.get("a-hyrmo"),   # int year*100+month (Tishrei=01), e.g. 578601
            "shana":     d.get("a-shana",   ""),
            "chodesh":   d.get("a-chodesh", ""),
            "shabbatot": shabbatot,
            "lastmo":    lastmo,
            "YomText":   str(d.get("YomText") or "").strip(),
            "ThisRC":    str(d.get("ThisRC")  or "").strip(),
            "NextRC":    str(d.get("NextRC")  or "").strip(),
        })
    wb.close()
    return months

# ── layout constants (in cm, from PDF analysis) ───────────────────────────────
# Shabbat table: 10 columns left→right (visual RTL order: motzash on far left)
# Widths measured from original PDF x-positions, scaled to BODY_W=27.7cm
SHAB_COLS = [
    # (top_label, bottom_label,  width_cm)
    ("מוצאי שבת",  "וערבית",         2.7),
    ("מנחה שבת",   "",               5.0),   # wide: holds "13:30  17:15  HH:MM"
    ("סוף זמן",    "קריאת שמע",      2.1),
    ("שחרית שבת",  "8:30",           2.1),
    ("שחרית שבת",  "6:45",           2.1),
    ("מנחה",       "וקבלת שבת",      2.2),
    ("הדלקת",      "נרות",           2.3),   # bold, candles — only ever a time (e.g. 18:42);
                                             # gave 0.8 of its width to תאריך עברי (see below)
    ("פרשת",       "השבוע",          4.0),
    ("תאריך",      "לועזי",          2.1),
    ("תאריך",      "עברי",           3.1),   # widened 2.3→3.1 so leap-year Adar dates
                                             # (e.g. 'י״ג אדר א׳') fit on ONE line. A wrap here
                                             # grows exact=False rows → 5-Shabbat month spills to
                                             # page 2. Width taken from הדלקת נרות, NOT from פרשת
                                             # (פרשת must fit long double-parshiot + "א' דראש השנה").
]
# Verify total ≈ BODY_W
_total = sum(w for _,_,w in SHAB_COLS)
# Scale to fit exactly
SHAB_COLS = [(l1, l2, w * BODY_W / _total) for l1, l2, w in SHAB_COLS]

# ── Theme (header cell fill + accent text) ──────────────────────────────────
# Selected via --theme; overridden at runtime in main(). Defaults to olive.
THEMES = {
    "olive":  {"hdr": "E6E9D8", "accent": "6B7040"},   # light olive fill, dark olive text
    "purple": {"hdr": "D9C7E8", "accent": "5B2D8C"},   # darker purple fill + text
    "gray":   {"hdr": "D8D8D8", "accent": "4472C4"},   # original gray + blue accent
}
THEME = THEMES["olive"]          # mutated by main() per --theme
def BG_HDR():   return THEME["hdr"]
def ACCENT():   return THEME["accent"]
BG_WHITE = "FFFFFF"

# Weekday section column widths (cm)
WK_L = 15.5   # מנחה/ערבית block width
WK_R = BODY_W - WK_L   # שחרית block width = 12.2 cm

# מנחה/ערבית inner columns
WK_ARVIT2  = 2.3
WK_DAF     = 2.3
WK_MINCHAV = 4.3
WK_SHAVUA  = WK_L - WK_ARVIT2 - WK_DAF - WK_MINCHAV  # ~6.6cm

# שחרית inner columns
WK_SH_YOM_AH = 2.6    # יום א-ה
WK_SH_FRI    = 3.2    # יום ו'
WK_SH_RC1    = (WK_R - WK_SH_YOM_AH - WK_SH_FRI) / 2  # two ראש חודש cols
WK_SH_RC2    = WK_SH_RC1

ORG_LINES = [
    ('"עמותת בית הכנסת יקיר אפרים במודיעין (ע"ר)"',                                                        9, True),
    ('ע"ש המנוח חיים אפרים פישל לבית בוימן ז"ל ורעיתו דבורה ז"ל אשר הונצח ע"י ההורים בוימן מלכה הי"ו ומשה ז"ל', 9, False),
    ('הרב דר. אנדי שיין, רחוב נחל שורק 27, מודיעין מיקוד 7170734, טל 08-9709264, yakirefraym@gmail.com',    9, False),
    ('Beit Knesset Yakir Efraim Modiin – Israel',                                                              9, False),
    ('Mailing Address – Rabbi Dr. Andrew Schein – 27, Nahal Soreq St. Modiin 7170734, Israel, Tel: 972-8-9709264', 9, False),
]
FOOTNOTE = '* מניין ראשון לתפילת ערבית מתקיים לאחר סיום תפילת מנחה (ולאחר השקיעה)'


# ── page builder ──────────────────────────────────────────────────────────────

def build_page(doc, month, logo_path, is_first, mincha_tbl=None, shabbat_tbl=None):
    if not is_first:
        _page_break(doc)

    mincha_tbl = mincha_tbl or {}
    shabbat_tbl = shabbat_tbl or {}
    shabbatot = month["shabbatot"]
    n = len(shabbatot)

    # Conditional notes for this month:
    chodesh = str(month.get("chodesh") or "").strip()
    is_kislev = ("כסלו" in chodesh)   # name-based (leap-year safe)
    # School-vacation note: Tishrei + Nissan always, plus any summer month that has a
    # Shabbat between 21 Jun and 31 Aug. Detect Tishrei/Nissan by NAME, not month
    # number: in a LEAP year Nissan is month 8 (7 = Adar II), so the old numeric
    # `heb_mo in (1,7)` put the note on Adar II and missed Nissan. (HCV bug: אדר ב׳
    # page wrongly showed בחופשות בתי הספר.)
    is_tishrei_or_nissan = ("תשרי" in chodesh) or ("ניסן" in chodesh)
    _dates = [sb.get("date") for sb in shabbatot
              if isinstance(sb.get("date"), (datetime.date, datetime.datetime))]
    _summer = any((d.month, d.day) >= (6, 21) and (d.month, d.day) <= (8, 31) for d in _dates)
    show_school = is_tishrei_or_nissan or _summer

    # ── HEADER (logo | title | בס״ד) ──────────────────────────────────────────
    ht = doc.add_table(rows=1, cols=3)
    _tbl_no_borders(ht)
    _tbl_w(ht._tbl, BODY_W)
    ht.alignment = WD_TABLE_ALIGNMENT.CENTER
    _row_h(ht.rows[0], 2.9)

    # logo (size from reference PDF: 132×81 pt ≈ 4.65 × 2.86 cm)
    lc = ht.rows[0].cells[0]; _cell_w(lc, 4.9)
    lc.text = ""
    lp = lc.paragraphs[0]; lp.alignment = WD_ALIGN_PARAGRAPH.LEFT; _spacing(lp, 0, 0)
    lc._tc.get_or_add_tcPr().append(_el("w:vAlign", **{"w:val": "center"}))
    if os.path.exists(logo_path):
        lp.add_run().add_picture(logo_path, width=Cm(4.65), height=Cm(2.86))

    # title
    tc = ht.rows[0].cells[1]
    tc.text = ""; tp = tc.paragraphs[0]
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER; _spacing(tp, 0, 0); _bidi_para(tp)
    tc._tc.get_or_add_tcPr().append(_el("w:vAlign", **{"w:val": "center"}))
    _run(tp, f'לו״ז חודש {month["chodesh"]} {month["shana"]}', F_ARIAL, 28, bold=True)

    # בס״ד
    bc = ht.rows[0].cells[2]; _cell_w(bc, 2.5)
    bc.text = ""; bp = bc.paragraphs[0]
    bp.alignment = WD_ALIGN_PARAGRAPH.RIGHT; _spacing(bp, 0, 0); _bidi_para(bp)
    bc._tc.get_or_add_tcPr().append(_el("w:vAlign", **{"w:val": "center"}))
    _run(bp, "בס״ד", F_ARIAL, 16, bold=True)

    # ── "זמני השבת" ────────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER; _bidi_para(p); _spacing(p, 3, 2)
    _run(p, "זמני השבת", F_ARIAL, 22, bold=True)

    # ── SHABBAT TABLE ──────────────────────────────────────────────────────────
    st = doc.add_table(rows=1 + n, cols=len(SHAB_COLS))
    _tbl_w(st._tbl, BODY_W)
    st.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Column widths on every row
    for row in st.rows:
        for ci, (_, _, wcm) in enumerate(SHAB_COLS):
            _cell_w(row.cells[ci], wcm)

    # Header: single tall row. Each column's two label lines (l1, l2) go into the
    # one header cell (l2 on a second line when present), vertically centered.
    # Exception: the two שחרית שבת columns (idx 3,4 = 8:30 / 6:45) are MERGED into
    # one header cell labelled "שחרית שבת" (matching the reference PDF); the data
    # rows below keep their two separate value columns.
    SHACH_COL_A, SHACH_COL_B = 3, 4
    _row_h(st.rows[0], 1.5, exact=True)
    hdr_cells = st.rows[0].cells
    merged = hdr_cells[SHACH_COL_A].merge(hdr_cells[SHACH_COL_B])
    for ci, (l1, l2, _) in enumerate(SHAB_COLS):
        if ci == SHACH_COL_A:
            _cell_write(merged, "שחרית שבת", F_CALIBRI, 16, before=1, after=1, line=18)
            _cell_bg(merged, BG_HDR())
            continue
        if ci == SHACH_COL_B:
            continue   # absorbed into the merged cell
        cell = hdr_cells[ci]
        _cell_write(cell, l1, F_CALIBRI, 16, before=0, after=0, line=18)
        if l2:
            _cell_para(cell, l2, F_CALIBRI, 16, before=0, after=0)
        _cell_bg(cell, BG_HDR())

    # Data rows
    for ri, sb in enumerate(shabbatot):
        dr = st.rows[1 + ri]
        _row_h(dr, 0.66, exact=False)
        # [motzash, min2, shma, shach2, shach1, frimin, candles, parsha, date_g, date_h]

        # CSV Shabbat block (keyed by this Shabbat's own date). Falls back to XLS.
        blk = {}
        dv = sb.get("date")
        if isinstance(dv, (datetime.date, datetime.datetime)):
            blk = shabbat_tbl.get(dv.strftime("%d/%m/%Y"), {})

        shach1_val = ft(sb.get("shach1"))
        if shach1_val in ("--", "---"): shach1_val = ""
        if isinstance(sb.get("shach1"), str) and sb.get("shach1","").strip() in ("--","---"):
            shach1_val = ""
        if blk.get("shacharit1"):        # CSV shacharit1 (default 6:45 or override)
            shach1_val = blk["shacharit1"]

        # Festival classification for this Shabbat (drives shach1 / 17:15 / blanking).
        d0 = None
        is_sukkot = is_no_minyan = False
        if isinstance(dv, (datetime.date, datetime.datetime)):
            d0 = dv.date() if isinstance(dv, datetime.datetime) else dv
            is_sukkot    = sukkot_shabbat(d0)          # Tishrei 15-21 (Kohelet → early)
            is_no_minyan = yomtov_no_minyan_shabbat(d0)  # Sim.Torah / Pesach 1st day

        # Shabbat Chol HaMoed shach1 override: Sukkot → 6:15, Pesach → 6:30 (takes
        # precedence). Extended: ANY Sukkot Shabbat (incl. 1st day, Tishrei 15) → 6:15,
        # because Kohelet is read and is long.
        if d0 is not None:
            chm = chol_hamoed_shabbat(d0)
            if chm:
                shach1_val = chm["shach1"]
            elif is_sukkot:
                shach1_val = "6:15"

        date_str = ""
        if isinstance(dv, (datetime.date, datetime.datetime)):
            # Israeli/European order, day first: 'D.M' (e.g. 15.8) — matches the
            # events list format and, being short (no leading zeros, dot not slash),
            # fits the narrow תאריך לועזי column on one line (a slashed '15/08'
            # wrapped to a stacked 15 / 08, which read ambiguously).
            date_str = f"{dv.day}.{dv.month}"
        elif dv:
            date_str = str(dv)[:5]

        # min2 (Shabbat mincha string), shma, friMin, candles, motzash:
        # prefer CSV block, fall back to XLS.
        # No Sharbat 17:15 on Sukkot / Shemini-Atzeret-Simchat-Torah Shabbatot.
        _suppress_17 = is_sukkot or is_no_minyan
        min2_val    = format_shabbat_mincha(blk, suppress_sharbat=_suppress_17) \
                      if blk.get("mincha_ktana") else str(sb.get("min2") or "").strip()
        shma_val    = blk.get("shma")     or ft(sb.get("shma"))
        frimin_val  = blk.get("mincha_erev") or ft(sb.get("friMin"))
        candles_val = blk.get("candles")  or ft(sb.get("candles"))
        motzash_val = blk.get("motzash")  or ft(sb.get("motzsh"))

        shach2_val = str(sb.get("shach2") or "").strip() or "8:30"   # constant
        # שמחת תורה / Pesach 1st day on Shabbat: the regular minyan schedule does not
        # apply — leave BOTH shacharit times and the Shabbat mincha string BLANK for
        # manual fill (special festival timetable).
        if is_no_minyan:
            shach1_val = ""
            shach2_val = ""
            min2_val   = ""

        vals = [
            motzash_val,
            min2_val,                            # "13:30  17:15  HH:MM" / "13:30  HH:MM"
            shma_val,
            shach2_val,                          # constant 8:30 (blank on no-minyan yom-tov)
            shach1_val,
            frimin_val,
            candles_val,                         # FROM PDF – never calculated
            str(sb.get("parsha") or "").strip(),
            date_str,
            blk.get("taarich") or str(sb.get("taarich") or "").strip(),
        ]
        # Four-parshiot tag (שקלים/זכור/פרה/החודש) for the parsha cell. Skip on
        # Chol HaMoed (parsha already relabeled שבת חוה"מ …; never one of the four).
        sp_tag = ""
        if isinstance(dv, (datetime.date, datetime.datetime)):
            d0 = dv.date() if isinstance(dv, datetime.datetime) else dv
            if not chol_hamoed_shabbat(d0):
                sp_tag = special_shabbat_tag(d0)

        for ci, v in enumerate(vals):
            # Bold שחרית שבת 1 (col 4) only when it differs from the usual 6:45,
            # to flag the exception. Candles (col 6) is no longer bold.
            bold = (ci == 4 and v.strip() not in ("", SHACHARIT_1_DEFAULT))
            _cell_write(dr.cells[ci], v, F_CALIBRI, 16, bold=bold,
                        before=0, after=0, line=16)
            # Parsha col (7): append the special-Shabbat tag as a smaller run,
            # same paragraph: "פרשה - שקלים".
            if ci == 7 and sp_tag:
                _run(dr.cells[ci].paragraphs[0], f"  -  {sp_tag}", F_CALIBRI, 12)
            _cell_bg(dr.cells[ci], BG_WHITE)

    _tbl_borders(st)

    # ── "תפילות יום חול" ───────────────────────────────────────────────────────
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER; _bidi_para(p2); _spacing(p2, 6, 0)
    _run(p2, "תפילות יום חול", F_ARIAL, 20, bold=True)

    # ── WEEKDAY OUTER TABLE (1 row × 2 cols: left=מנחה, right=שחרית) ──────────
    outer = doc.add_table(rows=1, cols=2)
    _tbl_no_borders(outer)
    _tbl_w(outer._tbl, BODY_W)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER

    left_cell  = outer.rows[0].cells[0]   # מנחה/ערבית
    right_cell = outer.rows[0].cells[1]   # שחרית

    _cell_w(left_cell,  WK_L)
    _cell_w(right_cell, WK_R)
    left_cell._tc.get_or_add_tcPr().append(_el("w:vAlign",  **{"w:val": "top"}))
    right_cell._tc.get_or_add_tcPr().append(_el("w:vAlign", **{"w:val": "top"}))

    # ── LEFT: מנחה/ערבית sub-table ────────────────────────────────────────────
    left_cell.text = ""
    lp0 = left_cell.paragraphs[0]; lp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _spacing(lp0, 0, 0); _bidi_para(lp0)
    # NOTE: if reverting the מנחה table to per-row daf/arvit constants, restore this
    # to "מנחה, דף יומי וערבית".
    _run(lp0, "מנחה וערבית", F_ARIAL, 15, bold=True)

    # Build week rows for מנחה table.
    # wkMin on row N = the correct mincha for row N's week label. No shift needed.
    # lastmo is NOT shown as a labeled row.
    # Weekday mincha source: modiin_weeks.csv (adj|calc), keyed by the week's
    # Sunday = Shabbat date + 1 day. Falls back to XLS wkMin if not found.
    wk_rows = []
    for w in shabbatot:
        lbl   = str(w.get("wkLbl") or "").strip()
        if not lbl: continue
        wmin = ""
        dv = w.get("date")
        if isinstance(dv, (datetime.date, datetime.datetime)):
            sunday = dv + datetime.timedelta(days=1)   # week after this Shabbat
            key = sunday.strftime("%d/%m/%Y")
            wmin = mincha_tbl.get(key, "")
            if not wmin:
                # No CSV row for that week (e.g. year boundary) — compute via ephem
                # with the same rule as modiin_weeks.py, so it's the true value.
                d0 = sunday.date() if isinstance(sunday, datetime.datetime) else sunday
                wmin = weekday_mincha_for_week(d0)
        if not wmin:                        # last-resort fallback: XLS-computed value
            wmin = ft(w.get("wkMin"), "")
        wk_rows.append((lbl, wmin))

    # Option A: 2-column table (שבוע | מנחה,ערבית1). daf(19:45)/arvit2(20:30) are
    # constant all year → stated once in a note below, not repeated every row.
    # RTL (bidiVisual): col0 = שבוע (rightmost), col1 = מנחה,ערבית1.
    # Narrow table (was full left-block width): שבוע sized to fit week labels, not
    # the whole column — leaves more empty space between this and the שחרית table.
    # מנחה column wide enough for the header 'מנחה וערבית1*' on ONE line (~4.1cm needed).
    # Widened from the left cell's idle slack (cell is 15.5cm, table only 4.5+9.0=13.5cm),
    # NOT by shrinking שבוע — so week labels stay safe.
    WK_M_MINCHA = 4.5
    # שבוע wide enough for the longest label — two double-parshiot joined, e.g.
    # 'תזריע-מצורע - אחרי מות-קדושים' (29 chars) — on ONE line (else the row wraps
    # and the month spills to a 2nd page).
    WK_M_SHAVUA = 9.0
    wkt_w = WK_M_MINCHA + WK_M_SHAVUA
    wkt = left_cell.add_table(rows=1 + len(wk_rows), cols=2)
    _tbl_w(wkt._tbl, wkt_w)
    _tbl_bidi(wkt)
    # The outer weekday table is RTL, so table alignment flips (same bidi quirk as
    # paragraph jc): to place this sub-table flush LEFT visually, set RIGHT here.
    # Gap then sits on its right (toward the center / שחרית table).
    wkt.alignment = WD_TABLE_ALIGNMENT.RIGHT
    _row_h(wkt.rows[0], 0.7, exact=True)
    for ci, (lbl, w) in enumerate([("שבוע", WK_M_SHAVUA), ("מנחה וערבית*", WK_M_MINCHA)]):
        _cell_write(wkt.rows[0].cells[ci], lbl, F_CALIBRI, 16, before=1, after=1)
        _cell_bg(wkt.rows[0].cells[ci], BG_HDR())
        _cell_w(wkt.rows[0].cells[ci], w)

    for ri, (lbl, wmin) in enumerate(wk_rows):
        dr = wkt.rows[1 + ri]
        _row_h(dr, 0.72, exact=False)
        for ci, (v, w) in enumerate([(lbl, WK_M_SHAVUA), (wmin, WK_M_MINCHA)]):
            _cell_write(dr.cells[ci], v, F_CALIBRI, 16,
                        align=WD_ALIGN_PARAGRAPH.CENTER,
                        before=0, after=0, line=18)
            _cell_w(dr.cells[ci], w)
    _tbl_borders(wkt)

    # ── RIGHT: שחרית sub-table ────────────────────────────────────────────────
    right_cell.text = ""
    rp0 = right_cell.paragraphs[0]; rp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _spacing(rp0, 0, 0); _bidi_para(rp0)
    _run(rp0, "שחרית", F_ARIAL, 15, bold=True)

    # שחרית weekday table — 4 columns of CONSTANT times (from reference PDF):
    # regular week + Rosh Chodesh variants. Header line 1 is a 2-col group span
    # ("שחרית" / "ראש חודש"); line 2 = יום א-ה / יום שישי under each group.
    # RTL (bidiVisual): col0 = rightmost.
    #   col0 = יום א-ה (regular)   col1 = יום שישי (regular)
    #   col2 = יום א-ה (ר"ח)       col3 = יום שישי (ר"ח)
    sh_cols = [
        ("יום א-ה",  SHACH_WK_YOM_AH),
        ("יום שישי", SHACH_WK_FRI),
        ("יום א-ה",  SHACH_WK_RC_YOM_AH),
        ("יום שישי", SHACH_WK_RC_FRI),
    ]
    sh_w = (WK_R - 0.1) / 4.0
    sht = right_cell.add_table(rows=3, cols=4)   # 2 header lines + 1 data line
    _tbl_w(sht._tbl, WK_R - 0.1)
    _tbl_bidi(sht)

    # Header line 1: group labels — "שחרית" over cols 0-1, "ראש חודש" over cols 2-3.
    _row_h(sht.rows[0], 0.7, exact=True)
    g1 = sht.rows[0].cells[0].merge(sht.rows[0].cells[1])
    _cell_write(g1, "שחרית", F_CALIBRI, 16, before=1, after=1); _cell_bg(g1, BG_HDR())
    g2 = sht.rows[0].cells[2].merge(sht.rows[0].cells[3])
    _cell_write(g2, "ראש חודש", F_CALIBRI, 16, before=1, after=1); _cell_bg(g2, BG_HDR())

    # Header line 2: day sub-labels
    _row_h(sht.rows[1], 0.7, exact=True)
    for ci, (lbl, _) in enumerate(sh_cols):
        _cell_write(sht.rows[1].cells[ci], lbl, F_CALIBRI, 16, before=1, after=1)
        _cell_bg(sht.rows[1].cells[ci], BG_HDR())
        _cell_w(sht.rows[1].cells[ci], sh_w)

    # Data line: the constant times
    _row_h(sht.rows[2], 0.72, exact=False)
    for ci, (_, val) in enumerate(sh_cols):
        _cell_write(sht.rows[2].cells[ci], val, F_CALIBRI, 16, before=0, after=0, line=18)
        _cell_w(sht.rows[2].cells[ci], sh_w)

    _tbl_borders(sht)

    # School-vacation note — relates to the Shacharit table (extra weekday minyan
    # during school breaks). Right cell, directly under the table, above the events.
    # Shown on Tishrei/Nissan + summer months (see show_school). align=None → RTL
    # right alignment via bidi (explicit jc=right renders left).
    # REUSE the empty paragraph python-docx leaves after the shacharit table (a table
    # can't be a cell's last element) — avoids an empty line before the note.
    if show_school:
        ps = right_cell.paragraphs[-1]
        _bidi_para(ps); _spacing(ps, 0, 4)   # snug under the table it supports; gap AFTER
        _run(ps, SCHOOL_NOTE, F_ARIAL, 11)

    # ── Events (holidays + Rosh Chodesh) — INSIDE the שחרית (right) cell, below the
    # shacharit table. Rendered as a borderless single-column RTL table (bidiVisual),
    # which right-aligns reliably like the other tables — loose paragraphs did not.
    yom = month.get("YomText", "").strip()
    if yom:
        evs = [e.strip() for e in yom.split("|") if e.strip()]
        if evs:
            et = right_cell.add_table(rows=len(evs), cols=1)
            _tbl_no_borders(et)
            _tbl_w(et._tbl, WK_R - 0.2)
            _tbl_bidi(et)
            for ri, ev in enumerate(evs):
                _row_h(et.rows[ri], 0.5, exact=False)
                # align=None → no w:jc; RTL bidi paragraph aligns right naturally
                # (an explicit jc=right on a bidi paragraph renders LEFT in Word).
                _cell_write(et.rows[ri].cells[0], ev, F_CALIBRI, 12, bold=True,
                            align=None, before=0, after=0, line=15)
                # accent color on the run
                for r in et.rows[ri].cells[0].paragraphs[0].runs:
                    r._r.get_or_add_rPr().append(_el("w:color", **{"w:val": ACCENT()}))

    # ── LEFT cell footer: footnote + weekday daf/arvit constants ─────────────────
    # These belong with the מנחה column (the * refers to מנחה,ערבית1; daf/arvit are
    # weekday times). Placing them in the LEFT cell keeps them side-by-side with the
    # right cell's events — no more overlap under the taller column.
    # REUSE the empty paragraph python-docx leaves after the מנחה table (a table
    # can't be a cell's last element) — avoids an empty line before the footnote.
    pf = left_cell.paragraphs[-1]
    _bidi_para(pf); pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT; _spacing(pf, 0, 1)
    _run(pf, FOOTNOTE, F_ARIAL, 12)

    # Chanukah note — Kislev only, directly under the arvit footnote.
    if is_kislev:
        pchan = left_cell.add_paragraph()
        _bidi_para(pchan); pchan.alignment = WD_ALIGN_PARAGRAPH.RIGHT; _spacing(pchan, 0, 1)
        _run(pchan, CHANUKAH_NOTE, F_ARIAL, 12)

    pc = left_cell.add_paragraph()
    _bidi_para(pc); pc.alignment = WD_ALIGN_PARAGRAPH.RIGHT; _spacing(pc, 1, 1)
    _run(pc, f"דף יומי: {WEEKDAY_DAF}  |  ערבית  {WEEKDAY_ARVIT2}  (בכל ימות השבוע)",
         F_CALIBRI, 18, bold=True)

    # ── Org info (footer box) ───────────────────────────────────────────────────
    # Blank spacer line(s) before the org block: 2 for a 4-shabbat month, 1 for a
    # 5-shabbat month (which is tighter). They're real empty paragraphs so you can
    # delete one by hand if a month ever spills to a 2nd page. Tight exact line
    # height keeps the org block itself compact.
    n_blank = 1 if n >= 5 else 2
    for _ in range(n_blank):
        bp = doc.add_paragraph()
        _spacing(bp, 0, 0, line=10)   # ~one blank text line
    for i, (txt, sz, bold) in enumerate(ORG_LINES):
        po = doc.add_paragraph()
        po.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _spacing(po, 0, 0, line=sz + 1)
        if any("֐" <= c <= "׿" for c in txt):
            _bidi_para(po)
        _run(po, txt, F_ARIAL, sz, bold=bold, rtl=bool(any("֐" <= c <= "׿" for c in txt)))


# ── document setup ─────────────────────────────────────────────────────────────

def make_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width    = Cm(PG_W)
    sec.page_height   = Cm(PG_H)
    sec.left_margin   = Cm(MAR)
    sec.right_margin  = Cm(MAR)
    sec.top_margin    = Cm(MAR_V)
    sec.bottom_margin = Cm(MAR_V)
    # Zero out default Normal style spacing
    ns = doc.styles["Normal"]
    ns.font.name = F_CALIBRI
    ns.font.size = Pt(12)
    ns.paragraph_format.space_before = Pt(0)
    ns.paragraph_format.space_after  = Pt(0)
    # Remove the blank paragraph Word inserts
    for p in list(doc.paragraphs):
        if not p.text.strip():
            p._element.getparent().remove(p._element)
            break
    return doc


# ── PDF conversion ─────────────────────────────────────────────────────────────

def to_pdf(docx_path):
    docx_path = os.path.abspath(docx_path)          # resolve relative paths first
    out_dir  = os.path.dirname(docx_path)
    pdf_path = docx_path.replace(".docx", ".pdf")
    for lo in ["libreoffice","libreoffice7.6","libreoffice7.5","soffice"]:
        if shutil.which(lo):
            r = subprocess.run([lo,"--headless","--convert-to","pdf",
                                "--outdir",out_dir,docx_path],
                               capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and os.path.exists(pdf_path): return pdf_path
    # Windows fallback: drive Word via PowerShell COM. Convert the /mnt/<drive>/
    # WSL path to a Windows path (e.g. /mnt/c/foo → C:\foo).
    try:
        def to_win(p):
            if p.startswith("/mnt/") and len(p) > 6 and p[6] == "/":
                return p[5].upper() + ":" + p[6:].replace("/", "\\")
            return p.replace("/", "\\")
        win, wp = to_win(docx_path), to_win(pdf_path)
        ps = ("$w=New-Object -ComObject Word.Application; $w.Visible=$false; "
              f"$d=$w.Documents.Open('{win}'); "
              f"$d.SaveAs([ref]'{wp}',[ref]17); $d.Close(); $w.Quit()")
        subprocess.run(["powershell.exe","-Command",ps],
                       capture_output=True, text=True, timeout=180)
        if os.path.exists(pdf_path): return pdf_path
    except Exception: pass
    return None


# ── main ───────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Yakir monthly calendar generator")
    ap.add_argument("--xls",    default=DEFAULT_XLS)
    ap.add_argument("--weeks-csv", default=DEFAULT_WEEKS_CSV,
                    help="modiin_weeks.csv — source for weekday mincha (adj|calc). "
                         "If missing, falls back to XLS-computed wkMin.")
    ap.add_argument("--logo",   default=DEFAULT_LOGO)
    ap.add_argument("--out",    default=None,
                    help="Output .docx. If omitted, writes yakir_calendar_output.docx next to "
                         "the --weeks-csv input.")
    ap.add_argument("--year", type=int, default=None,
                    help="Hebrew year, e.g. 5787. Generates ALL its months in order (12, or 13 "
                         "in a leap year — figured out automatically from the CSV). The easy "
                         "way to make a full-year calendar; no need to list hyrmo keys.")
    ap.add_argument("--months", type=int, nargs="+", default=None,
                    help="0-based month indices into the sheet (fragile — prefer --hyrmo). "
                         "Default: all months.")
    ap.add_argument("--hyrmo", type=int, nargs="+", default=None,
                    help="Select individual months by year+month key (Tishrei=01), e.g. 578601 "
                         "578602 = Tishrei & Cheshvan 5786. For a whole year, prefer --year.")
    ap.add_argument("--source", choices=["auto", "csv", "xls"], default="auto",
                    help="Month-structure source. 'auto' (default) uses modiin_weeks.csv if present, "
                         "else the XLS. 'csv' forces CSV; 'xls' forces the XLS.")
    ap.add_argument("--theme", choices=list(THEMES.keys()), default="olive",
                    help="Color scheme for header cells + event accent (default: olive).")
    ap.add_argument("--pdf",    action="store_true")
    args = ap.parse_args()

    global THEME
    THEME = THEMES[args.theme]

    # Default output = yakir_calendar_output.docx next to the --weeks-csv input.
    out_path = args.out or os.path.join(os.path.dirname(os.path.abspath(args.weeks_csv)),
                                        "yakir_calendar_output.docx")

    # Choose the month-structure source.
    use_csv = (args.source == "csv") or (args.source == "auto" and os.path.exists(args.weeks_csv))
    if use_csv:
        if not os.path.exists(args.weeks_csv):
            print(f"ERROR: --source csv but not found: {args.weeks_csv}", file=sys.stderr); sys.exit(1)
        print(f"Loading months from {os.path.basename(args.weeks_csv)}…")
        all_months = load_months_from_csv(args.weeks_csv)
    else:
        if not os.path.exists(args.xls):
            print(f"ERROR: XLS not found: {args.xls}", file=sys.stderr); sys.exit(1)
        print("Loading months from XLS…")
        all_months = load_months(args.xls)
    print(f"  {len(all_months)} months")

    mincha_tbl = load_weekday_mincha(args.weeks_csv)
    if mincha_tbl:
        print(f"  weekday mincha: {len(mincha_tbl)} week(s) from {os.path.basename(args.weeks_csv)}")
    else:
        print(f"  weekday mincha: modiin_weeks.csv not found — falling back to XLS wkMin")

    shabbat_tbl = load_shabbat_block(args.weeks_csv)
    if shabbat_tbl:
        print(f"  shabbat block: {len(shabbat_tbl)} shabbat(ot) from {os.path.basename(args.weeks_csv)}")
    else:
        print(f"  shabbat block: modiin_weeks.csv not found — falling back to XLS")

    if args.year:
        # Select every month of the Hebrew year, in calendar order (Tishrei→Elul).
        # Leap years just work: 13 months exist in the CSV, so 13 come through.
        months = sorted((m for m in all_months
                         if str(m.get("hyrmo")).strip().isdigit()
                         and int(str(m.get("hyrmo")).strip()) // 100 == args.year),
                        key=lambda m: int(str(m.get("hyrmo")).strip()))
        if not months:
            print(f"  ⚠ no months found for year {args.year} in the CSV. "
                  f"Check --weeks-csv covers that year.", file=sys.stderr)
        elif len(months) not in (12, 13):
            print(f"  ⚠ year {args.year}: found {len(months)} month(s) (expected 12, or 13 in a "
                  f"leap year). The CSV may not fully cover this year.", file=sys.stderr)
    elif args.hyrmo:
        by_key = {str(m.get("hyrmo")).strip(): m for m in all_months}
        months = []
        for k in args.hyrmo:
            m = by_key.get(str(k).strip())
            if m:
                months.append(m)
            else:
                print(f"  ⚠ hyrmo {k} not found — skipped", file=sys.stderr)
    elif args.months:
        months = [all_months[i] for i in args.months if 0 <= i < len(all_months)]
    else:
        months = all_months
    print(f"  Generating {len(months)} month(s)")

    doc = make_doc()
    for idx, m in enumerate(months):
        print(f"  {idx+1:2d}/{len(months)}: {m['chodesh']} {m['shana']} ({len(m['shabbatot'])} shabbatot)")
        build_page(doc, m, args.logo, is_first=(idx == 0),
                   mincha_tbl=mincha_tbl, shabbat_tbl=shabbat_tbl)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    print(f"\nSaved: {out_path}")

    if args.pdf:
        print("Converting to PDF…")
        pdf = to_pdf(out_path)
        print(f"PDF: {pdf}" if pdf else "PDF failed — use Word: File → Save As → PDF")
    print("Done.")


if __name__ == "__main__":
    main()
