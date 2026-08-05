#!/usr/bin/env python3
"""
modiin_weeks.py — build modiin_weeks.csv (one row per Sun–Sat week).

Reads modiin_calendar.csv (produced by parse_modiin_pdf.py) and emits a
week-oriented data file that carries both PDF-sourced weekend facts and
ephem-computed weekday sunset data.

Each week is anchored on its Saturday (a Saturday row in modiin_calendar.csv).
  From = Saturday − 6 days (Sunday)
  To   = Saturday − 2 days (Thursday)
Weekday sunset span = Sun..Thu (5 days), computed with ephem.

ephem sunset is calibrated against official Modiin shkia (see calibrate_ephem.py):
  horizon -1:00, pressure=0, at Modiin center → mean bias -0.02 min, MAE 0.70 min.

WHY THIS FILE EXISTS: the master XLS's daily sunset data ends 2030-12-31. This
CSV pipeline (ephem-based, no end date) is replacing the XLS as the source of
truth for the Word/PDF calendar. The XLS is now validation-only; it is NOT read
here. Columns are ported from the XLS one rule at a time and validated against it.

Columns emitted (modiin_weeks.csv):
  week                     Sequential week number (1..N).
  hyrmo                    Sort/lookup key = heb_year*100 + heb_month_num (Tishrei=01),
                           e.g. 578602 = Cheshvan 5786. Matches the XLS a-hyrmo convention.
  heb_year                 Hebrew year as an integer, e.g. 5786.
  heb_year_str             Hebrew year in letters, e.g. תשפ״ו.
  heb_month_num            Hebrew month number, Tishrei=1 (handles leap-year Adar I/II).
  heb_month_name           Hebrew month name, e.g. חשון / אדר א׳ (from pyluach, matches XLS).
  from                     Sunday date of the week, DD/MM/YYYY (Saturday − 6).
  to                       Thursday date of the week, DD/MM/YYYY (Saturday − 2).
  weekend_name             Parasha name for that week's Shabbat (from modiin_calendar.csv).
  weekend_date             The anchoring Saturday date, DD/MM/YYYY.
  --- Shabbat block ---
  candles                  Candle-lighting, from the Modiin PDF — never computed.
  motzash                  Motzaei Shabbat, from the Modiin PDF — never computed.
  sharbat_mincha           yes/no — summer 17:15 mincha (organized by Yair Sharbat).
                           "yes" if candles > 18:00.
  shabbat_mincha_ktana_calc  COMPUTED late Shabbat mincha = floor_to_5(candles).
  shabbat_mincha_ktana_adj   MANUAL override; wins over _calc.
  shacharit1               Shacharit-1: blank → default 6:45; fill to override (e.g. RC 6:30).
  sof_zman_shma_calc       COMPUTED sof zman shma (GRA) via ephem, calibrated (MAE 0.36 min).
  sof_zman_shma_adj        MANUAL override; wins over _calc.

  NOTE: Mincha erev shabbat (Friday) is NOT a column. It is derived by the Word
  generator as (resolved mincha ktana) + 10 min, matching the old calendar exactly
  (verified 50/50 weeks). A ktana override therefore shifts erev automatically.
  --- weekday block ---
  sunset_rabanut           Official Saturday sunset from the Modiin PDF (shkia).
  sunset_ephem             ephem-computed Saturday sunset, HH:MM.
  sunset_delta             ephem − rabanut, integer minutes (sanity check, should be ≈0).
  earliest_sunset_weekday  Earliest ephem sunset among Sun–Thu, HH:MM. Basis for mincha.
  weekday_sunset_delta     Spread = latest − earliest weekday sunset, integer minutes.
  weekday_mincha_calc      COMPUTED weekday mincha = floor_to_5(earliest_sunset_weekday − 12 min).
  weekday_mincha_adj       MANUAL override, blank by default. When non-blank it wins over _calc.

Shabbat-block constants (NOT columns — see RULES.md §5):
  mincha gedola 13:30 (always), Sharbat mincha 17:15 (when sharbat_mincha=yes),
  shacharit 2 = 8:30. Holiday exceptions are NOT handled (dedicated holiday timetables).

Hebrew month/year:
  Computed from the week's Shabbat (weekend_date) via pyluach. Month is numbered from
  Tishrei (Tishrei=1), matching the XLS. This lets the calendar be selected by year+month
  (hyrmo) instead of a fragile 0-based sheet index.

Weekday mincha rule:
  mincha = floor_to_5( earliest Sun–Thu sunset − WEEKDAY_MINCHA_OFFSET_MIN )
  Offset is 12 min (was 15 in the old XLS "Andy Mincha" column). Basing it on the
  EARLIEST sunset of the week guarantees the service finishes before sunset every
  weekday.

calc vs adj contract:
  _calc is always recomputed from sunset data on every run.
  _adj is a human-editable override. Consumers (e.g. yakir_calendar.py) use
  `adj if adj else calc`. To protect manual edits, regeneration carries forward any
  existing _adj values from the current CSV (keyed by weekend_date) unless
  --no-preserve-adj is passed.
"""
import argparse
import csv
import datetime
import os
import sys
from zoneinfo import ZoneInfo

try:
    import ephem
except ImportError:
    sys.exit("ephem not installed. Run: pip3 install ephem --break-system-packages")

try:
    from pyluach import dates as _hebdates, hebrewcal as _hebcal
except ImportError:
    sys.exit("pyluach not installed. Run: pip3 install pyluach --break-system-packages")

# ---- Calibrated constants (from calibrate_ephem.py) ------------------------
LAT = 31.8981          # Modiin center
LON = 35.0104
TZ = ZoneInfo("Asia/Jerusalem")
DEFAULT_HORIZON = "-1:00"   # best fit vs official Modiin shkia (MAE 0.70 min)
SHMA_HORIZON = "-0:50"      # best fit vs XLS OU GRA shma (MAE 0.36 min) — calibrated
                            # independently of the sunset horizon; do not merge the two.

# ---- Shabbat prayer-time constants (see RULES.md §5) -----------------------
SHABBAT_MINCHA_GEDOLA = "13:30"   # constant, always shown
SHARBAT_MINCHA        = "17:15"   # constant summer mincha (Yair Sharbat), shown when sharbat_mincha=yes
SHACHARIT_2           = "8:30"    # constant 2nd minyan (holiday exceptions handled separately)
SHACHARIT_1_DEFAULT   = "6:45"    # default 1st minyan; shacharit1 column overrides when filled
# Note: mincha erev shabbat = mincha ktana + 10 min is applied by the Word generator
# (yakir_calendar.py) at render time, not stored here — so a ktana override shifts it.

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Inputs default to the CURRENT working directory (never the code folder), so running
# from a year folder "just works". Output goes next to the input --csv.
DEFAULT_CSV = os.path.join(os.getcwd(), "modiin_calendar.csv")

SUMMER_CANDLES_THRESHOLD = (18, 0)    # candles > 18:00 => Sharbat mincha (17:15)
WEEKDAY_MINCHA_OFFSET_MIN = 12        # mincha = floor_to_5(earliest weekday sunset − 12 min)

FIELDNAMES = [
    "week", "hyrmo",
    "heb_year", "heb_year_str", "heb_month_num", "heb_month_name",
    "from", "to",
    "weekend_name", "weekend_date",
    # ── Shabbat block ────────────────────────────────────────────────
    "taarich",                   # Hebrew date of the Shabbat, e.g. 'ה׳ תשרי'
    "candles",                   # from PDF (via modiin_calendar.csv) — never computed
    "motzash",                   # from PDF — never computed
    "sharbat_mincha",            # yes/no — summer 17:15 mincha (organized by Yair Sharbat)
    "shabbat_mincha_ktana_calc", # = floor_to_5(candles)
    "shabbat_mincha_ktana_adj",  # manual override; wins over _calc
    "shacharit1",                # blank → default 6:45; fill to override (e.g. Rosh Chodesh 6:30)
    "sof_zman_shma_calc",        # GRA: sunrise + (sunset−sunrise)/4  (ephem, calibrated)
    "sof_zman_shma_adj",         # manual override; wins over _calc
    # ── weekday block ────────────────────────────────────────────────
    "sunset_rabanut", "sunset_ephem", "sunset_delta",
    "earliest_sunset_weekday", "weekday_sunset_delta",
    "weekday_mincha_calc",       # computed: floor_to_5(earliest_sunset_weekday − 12)
    "weekday_mincha_adj",        # manual override, blank by default; wins over _calc when set
]


# ---- ephem sunset (calibrated) ---------------------------------------------
def ephem_sunset_minutes(date, horizon=DEFAULT_HORIZON):
    """Local sunset as float minutes-since-midnight for the given date."""
    obs = ephem.Observer()
    obs.lat = str(LAT)
    obs.lon = str(LON)
    obs.elevation = 0
    obs.horizon = str(horizon)
    obs.pressure = 0   # disable ephem's own refraction; horizon fully controls it
    noon_local = datetime.datetime(date.year, date.month, date.day, 12, 0, tzinfo=TZ)
    obs.date = ephem.Date(noon_local.astimezone(datetime.timezone.utc).replace(tzinfo=None))
    setting_utc = obs.next_setting(ephem.Sun()).datetime().replace(tzinfo=datetime.timezone.utc)
    local = setting_utc.astimezone(TZ)
    return local.hour * 60 + local.minute + local.second / 60.0


def ephem_sunrise_minutes(date, horizon=SHMA_HORIZON):
    """Local sunrise as float minutes-since-midnight for the given date."""
    obs = ephem.Observer()
    obs.lat = str(LAT)
    obs.lon = str(LON)
    obs.elevation = 0
    obs.horizon = str(horizon)
    obs.pressure = 0   # disable ephem's own refraction; horizon fully controls it
    noon_local = datetime.datetime(date.year, date.month, date.day, 12, 0, tzinfo=TZ)
    obs.date = ephem.Date(noon_local.astimezone(datetime.timezone.utc).replace(tzinfo=None))
    rising_utc = obs.previous_rising(ephem.Sun()).datetime().replace(tzinfo=datetime.timezone.utc)
    local = rising_utc.astimezone(TZ)
    return local.hour * 60 + local.minute + local.second / 60.0


def gra_shma_minutes(date, horizon=SHMA_HORIZON):
    """
    Sof zman kriat shma (GRA) as float minutes-since-midnight.
    GRA: shma = sunrise + 3 halachic hours, where a halachic hour = (sunset−sunrise)/12.
    So shma = sunrise + (sunset − sunrise) / 4.
    Calibrated to ±0.36 min MAE vs the XLS OU "Shema Latest, GRA" values.
    """
    sr = ephem_sunrise_minutes(date, horizon)
    ss = ephem_sunset_minutes(date, horizon)
    return sr + (ss - sr) / 4.0


def hhmm(total_minutes):
    """Round float minutes-since-midnight to nearest minute → 'HH:MM'."""
    m = int(round(total_minutes)) % (24 * 60)
    return f"{m // 60:02d}:{m % 60:02d}"


def floor5(minutes):
    """Floor an integer minutes-since-midnight value down to a multiple of 5."""
    return (int(minutes) // 5) * 5


def hebrew_month_fields(greg_date):
    """
    Given a Gregorian date, return the Hebrew calendar fields for the month it
    falls in, numbered from Tishrei (Tishrei=1), matching the XLS convention.

    Returns dict: heb_year (int, e.g. 5786), heb_year_str (letters, e.g. תשפ״ו),
    heb_month_num (int, 1=Tishrei .. handles leap-year Adar I/II), heb_month_name
    (Hebrew, e.g. חשון / אדר א׳), hyrmo (int sort key = heb_year*100 + heb_month_num).

    Month names, Hebrew year letters and leap-year Adar handling all come from
    pyluach so they stay consistent with the XLS strings.
    """
    h = _hebdates.GregorianDate(greg_date.year, greg_date.month, greg_date.day).to_heb()
    # Tishrei-based ordinal: itermonths() yields the year's months in civil order
    # (Tishrei first), so the 1-based position IS the Tishrei-based month number.
    civil = [m.month for m in _hebcal.Year(h.year).itermonths()]
    month_num = civil.index(h.month) + 1
    month_name = h.month_name(hebrew=True)
    return {
        "heb_year": h.year,
        "heb_year_str": _hebcal.Year(h.year).year_string(),
        "heb_month_num": month_num,
        "heb_month_name": month_name,
        "hyrmo": h.year * 100 + month_num,
        # Hebrew date for display, e.g. 'ה׳ תשרי' (day-of-month + month name)
        "taarich": f"{h.hebrew_day()} {month_name}",
    }


def parse_hhmm(s):
    """'HH:MM' → minutes-since-midnight (int), or None for blank/'-----'."""
    s = (s or "").strip()
    if not s or s.startswith("-"):
        return None
    hh, mm = s.split(":")
    return int(hh) * 60 + int(mm)


# ---- data loading ----------------------------------------------------------
def load_calendar_rows(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


# Manual override columns preserved across regeneration (keyed by weekend_date).
MANUAL_COLUMNS = [
    "shabbat_mincha_ktana_adj",
    "shacharit1",
    "sof_zman_shma_adj",
    "weekday_mincha_adj",
]


def load_existing_manual(path):
    """Read existing manual-override columns from a prior CSV, keyed by
    weekend_date → {column: value}. Returns {} if the file is missing.
    Only non-blank values for columns in MANUAL_COLUMNS are captured."""
    if not os.path.exists(path):
        return {}
    saved = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            key = (row.get("weekend_date") or "").strip()
            if not key:
                continue
            vals = {c: (row.get(c) or "").strip()
                    for c in MANUAL_COLUMNS if (row.get(c) or "").strip()}
            if vals:
                saved[key] = vals
    return saved


def build_weeks(rows, horizon=DEFAULT_HORIZON):
    weeks = []
    n = 0
    for r in rows:
        date_g = (r.get("date_g") or "").strip()
        try:
            sat = datetime.datetime.strptime(date_g, "%d/%m/%Y").date()
        except ValueError:
            continue
        if sat.weekday() != 5:   # 5 = Saturday
            continue

        n += 1
        sunday = sat - datetime.timedelta(days=6)
        thursday = sat - datetime.timedelta(days=2)

        # Sharbat mincha (17:15) rule: candles LATER than 18:00 => yes
        candles_min = parse_hhmm(r.get("candles"))
        thr = SUMMER_CANDLES_THRESHOLD[0] * 60 + SUMMER_CANDLES_THRESHOLD[1]
        sharbat_mincha = "yes" if (candles_min is not None and candles_min > thr) else "no"

        # Shabbat mincha ktana = floor_to_5(candles). Blank on holiday/no-candles rows.
        # Mincha erev shabbat (Friday) is NOT stored: it is derived at render time as
        # (resolved ktana) + 10 min, so a ktana override automatically shifts it.
        if candles_min is not None:
            mincha_ktana_calc = hhmm(floor5(candles_min))
        else:
            mincha_ktana_calc = ""

        # Sof zman shma (GRA) for this Shabbat, via ephem (calibrated).
        shma_calc = hhmm(gra_shma_minutes(sat))

        # Saturday sunset: official (rabanut) + ephem + delta
        rab_min = parse_hhmm(r.get("shkia"))
        sat_ephem_min = ephem_sunset_minutes(sat, horizon)
        sunset_delta = ""
        if rab_min is not None:
            sunset_delta = int(round(sat_ephem_min - rab_min))

        # Weekday sunsets Sun..Thu (5 days)
        wk_sunsets = [
            ephem_sunset_minutes(sunday + datetime.timedelta(days=i), horizon)
            for i in range(5)   # Sun, Mon, Tue, Wed, Thu
        ]
        earliest_min = int(round(min(wk_sunsets)))   # basis for weekday mincha
        weekday_delta = int(round(max(wk_sunsets) - min(wk_sunsets)))
        # Weekday mincha rule: floor_to_5(earliest weekday sunset − 12 min).
        mincha_calc = hhmm(floor5(earliest_min - WEEKDAY_MINCHA_OFFSET_MIN))

        # Hebrew month/year for this week, anchored on its Shabbat (sat).
        hm = hebrew_month_fields(sat)

        weeks.append({
            "week": n,
            "hyrmo": hm["hyrmo"],
            "heb_year": hm["heb_year"],
            "heb_year_str": hm["heb_year_str"],
            "heb_month_num": hm["heb_month_num"],
            "heb_month_name": hm["heb_month_name"],
            "from": sunday.strftime("%d/%m/%Y"),
            "to": thursday.strftime("%d/%m/%Y"),
            "weekend_name": (r.get("parsha") or "").strip(),
            "weekend_date": sat.strftime("%d/%m/%Y"),
            # Shabbat block
            "taarich": hm["taarich"],
            "candles": (r.get("candles") or "").strip(),
            "motzash": (r.get("motzash") or "").strip(),
            "sharbat_mincha": sharbat_mincha,
            "shabbat_mincha_ktana_calc": mincha_ktana_calc,
            "shabbat_mincha_ktana_adj": "",
            "shacharit1": "",                # blank → default 6:45 at render
            "sof_zman_shma_calc": shma_calc,
            "sof_zman_shma_adj": "",
            # weekday block
            "sunset_rabanut": (r.get("shkia") or "").strip(),
            "sunset_ephem": hhmm(sat_ephem_min),
            "sunset_delta": sunset_delta,
            "earliest_sunset_weekday": hhmm(earliest_min),
            "weekday_sunset_delta": weekday_delta,
            "weekday_mincha_calc": mincha_calc,
            "weekday_mincha_adj": "",        # manual override, blank by default
        })
    return weeks


def save_weeks_csv(weeks, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        w.writerows(weeks)


def main():
    p = argparse.ArgumentParser(description="Build modiin_weeks.csv from modiin_calendar.csv")
    p.add_argument("--csv", default=DEFAULT_CSV, help="Input modiin_calendar.csv")
    p.add_argument("--out", default=None,
                   help="Output modiin_weeks.csv. If omitted, writes next to the input --csv.")
    p.add_argument("--horizon", default=DEFAULT_HORIZON,
                   help="ephem horizon angle (calibration knob), e.g. -1:00")
    p.add_argument("--no-preserve-adj", action="store_true",
                   help="Do NOT carry forward existing weekday_mincha_adj overrides from --out "
                        "(default: preserve them). Use to reset all manual overrides to blank.")
    args = p.parse_args()

    if not os.path.exists(args.csv):
        sys.exit(f"Input CSV not found: {args.csv}")

    # Default output = modiin_weeks.csv next to the input calendar CSV.
    out_path = args.out or os.path.join(os.path.dirname(os.path.abspath(args.csv)),
                                        "modiin_weeks.csv")

    rows = load_calendar_rows(args.csv)
    weeks = build_weeks(rows, args.horizon)

    # Carry forward ALL manual-override columns from an existing output file,
    # unless explicitly disabled. Protects hand-edited values across regeneration.
    n_preserved = 0
    if not args.no_preserve_adj:
        saved = load_existing_manual(out_path)
        for w in weeks:
            vals = saved.get(w["weekend_date"])
            if vals:
                for col, v in vals.items():
                    w[col] = v
                n_preserved += 1

    save_weeks_csv(weeks, out_path)

    # Summary + delta sanity report
    deltas = [w["sunset_delta"] for w in weeks if isinstance(w["sunset_delta"], int)]
    print(f"Wrote {len(weeks)} week rows → {os.path.abspath(out_path)}")
    if n_preserved:
        print(f"  preserved manual overrides on {n_preserved} week(s)")
    if deltas:
        worst = max(deltas, key=abs)
        over2 = [w for w in weeks if isinstance(w["sunset_delta"], int) and abs(w["sunset_delta"]) > 2]
        print(f"sunset_delta (ephem − rabanut): "
              f"min {min(deltas)}, max {max(deltas)}, worst |{worst}| min")
        if over2:
            print(f"  ⚠ {len(over2)} week(s) off by >2 min:")
            for w in over2:
                print(f"    week {w['week']} {w['weekend_date']} {w['weekend_name']}: {w['sunset_delta']:+d} min")
        else:
            print("  ✓ all weeks within ±2 min of official shkia")


if __name__ == "__main__":
    main()
