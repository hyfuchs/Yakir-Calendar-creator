# Yakir Calendar Creator — Rules & Technical Reference

**Congregation**: Beit Knesset Yakir Efraim, Modiin  
**Purpose**: Generate monthly Hebrew-calendar prayer-time tables from official city times  
**Stack**: Python 3 · openpyxl · python-docx · pymupdf

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Data Flow](#2-data-flow)
3. [The Three Tools](#3-the-three-tools)
4. [Critical Rule: Candles & Motzash Are Sacred](#4-critical-rule-candles--motzash-are-sacred)
5. [Prayer Time Rules](#5-prayer-time-rules)
6. [XLS Architecture](#6-xls-architecture)
7. [Output Document Layout](#7-output-document-layout)
8. [Technical Appendix for AI](#8-technical-appendix-for-ai)

---

## 1. Project Overview

Each year, the congregation produces a 12-page A4-landscape Word document — one page per
Jewish month — showing all Shabbat and weekday prayer times.

**Old process**: Feed XLS → mail-merge in Word → manual editing  
**New process**: Python reads XLS → generates Word directly, no mail-merge

The annual Modiin city candles PDF is the official source of candle-lighting and Havdalah
times. All other prayer times are calculated by the XLS formulas.

---

## 2. Data Flow

```
Modiin City PDF
  (candles, shkia, motzash — Shabbat/Yom Tov only)
        │
        ▼
  parse_modiin_pdf.py
        │  outputs
        ▼
  modiin_calendar.csv
  (59 rows, one per Shabbat/Yom Tov)
        │
        ▼
  populate_xls.py
        │  writes raw inputs to Detail sheet
        ▼
  XLS Detail sheet
  (columns AJ=candles, AK=motzash, AL=shkia, AO=parsha, ...)
        │
        │  XLS formulas calculate all prayer times automatically
        ▼
  XLS ResultForWordMerge sheet
  (15 rows, one per Jewish month, all times pre-calculated)
        │
        ▼
  yakir_calendar.py
        │  reads ResultForWordMerge, builds Word doc
        ▼
  yakir_calendar_output.docx
  (one page per month, landscape A4, RTL Hebrew)
        │  optional
        ▼
  yakir_calendar_output.pdf
```

### Weekday מנחה/ערבית times — source and validity

Weekday mincha is **computed from ephem sunset data** in `modiin_weeks.py` and
written to `modiin_weeks.csv`. It is no longer taken from the XLS.

**Rule**: `weekday_mincha_calc = floor_to_5( earliest Sun–Thu sunset − 12 min )`.

- The **earliest** sunset of the weekday span (Sun–Thu) is the basis, so the
  service finishes before sunset on every weekday of the week.
- Offset is **12 minutes** (the old XLS "Andy Mincha" column used 15 min; the
  12-min rule lands 0–5 min later — identical to the old calendar on ~33 of 53
  weeks in 5786, +5 min on the rest).
- `floor_to_5` rounds **down** to a 5-minute boundary.

**Manual override**: `modiin_weeks.csv` also has a `weekday_mincha_adj` column,
blank by default. When a value is entered there it **wins** over `_calc`.
Consumers use `adj if adj else calc`. Regenerating the CSV preserves existing
`_adj` values (keyed by `weekend_date`) unless `--no-preserve-adj` is passed.

**Source of the value in the Word doc**: `yakir_calendar.py` reads
`modiin_weeks.csv` (`--weeks-csv`, default `modiin_weeks.csv` in the project folder) and looks up
the mincha by the week's **Sunday** date. If the CSV is missing it falls back to
the XLS `wkMin` (legacy behavior).

**Join/shift alignment**: the calendar shows, on Shabbat row **D**, the mincha
for the week that **starts the next day**. So the generator looks up key =
`D + 1 day` (that week's Sunday), which equals the `from` column of the
`modiin_weeks.csv` row. Example: Shabbat 25/10/2025 → key 26/10/2025 → 16:40.

**Validity**: ephem sunset has no end date (unlike the XLS `sunset` sheet, which
stopped at 2030-12-31). ephem is calibrated to ±0.7 min MAE vs official Modiin
shkia. The XLS is now used only to validate ported rules, not as a data source.

### modiin_weeks.csv columns

One row per week, anchored on its Saturday. Weekday span = preceding Sun–Thu
(`from` = Sat−6, `to` = Sat−2).

| Column | Description |
|---|---|
| `week` | Sequential week number (1..N) |
| `hyrmo` | Sort/lookup key = heb_year×100 + heb_month_num (Tishrei=01), e.g. 578602. Matches XLS `a-hyrmo` |
| `heb_year` | Hebrew year as integer, e.g. 5786 |
| `heb_year_str` | Hebrew year in letters, e.g. תשפ״ו |
| `heb_month_num` | Hebrew month number, Tishrei=1 (handles leap-year Adar I/II) |
| `heb_month_name` | Hebrew month name, e.g. חשון / אדר א׳ (from pyluach, matches XLS) |
| `from` | Sunday of the week, DD/MM/YYYY (Sat − 6) — join key for weekday mincha |
| `to` | Thursday of the week, DD/MM/YYYY (Sat − 2) |
| `weekend_name` | Parasha name for that week's Shabbat |
| `weekend_date` | The anchoring Saturday date, DD/MM/YYYY — also determines the Hebrew month |
| `taarich` | Hebrew date of the Shabbat, e.g. `ה׳ תשרי` (pyluach) |
| `candles` | Candle-lighting, from the Modiin PDF — never computed |
| `motzash` | Motzaei Shabbat, from the Modiin PDF — never computed |
| `sharbat_mincha` | yes/no — summer 17:15 mincha (Yair Sharbat); "yes" if candles **> 18:00** |
| `shabbat_mincha_ktana_calc` | COMPUTED late Shabbat mincha = floor_to_5(candles) |
| `shabbat_mincha_ktana_adj` | MANUAL override; wins over `_calc` |
| `shacharit1` | Shacharit-1: blank → default 6:45; fill to override (e.g. RC 6:30) |
| `sof_zman_shma_calc` | COMPUTED sof zman shma (GRA) via ephem, calibrated (MAE 0.36 min) |
| `sof_zman_shma_adj` | MANUAL override; wins over `_calc` |
| `sunset_rabanut` | Official Saturday sunset from Modiin PDF (shkia) |
| `sunset_ephem` | ephem-computed Saturday sunset, HH:MM |
| `sunset_delta` | ephem − rabanut, integer minutes (sanity check, ≈0) |
| `earliest_sunset_weekday` | Earliest ephem sunset among Sun–Thu — basis for weekday mincha |
| `weekday_sunset_delta` | Spread = latest − earliest weekday sunset, integer minutes |
| `weekday_mincha_calc` | COMPUTED mincha = floor_to_5(earliest_sunset_weekday − 12) |
| `weekday_mincha_adj` | MANUAL override, blank by default; wins over `_calc` when set |

**Hebrew month/year**: computed from each week's Shabbat (`weekend_date`) via `pyluach`, numbered
from **Tishrei = 1** (matching the XLS). The calendar can be selected by year+month using `hyrmo`
(`yakir_calendar.py --hyrmo 578601 578602` = Tishrei & Cheshvan 5786) instead of a fragile 0-based
sheet index.

**Migration note**: the project is moving off the XLS. `modiin_weeks.csv` /
`modiin_calendar.csv` are becoming the source of truth for the Word/PDF calendar.
The `calc` + `adj` pair is the pattern for future ported columns — a computed
value plus a blank manual-override column that takes precedence.

### Month structure — built from the CSV

`yakir_calendar.py` builds each month directly from `modiin_weeks.csv` (function
`load_months_from_csv`). Months are grouped by `hyrmo` in `week` order; the XLS is
no longer required. `--source auto` (default) uses the CSV when present, else the
XLS; `--source xls` forces the old path.

- **Which Shabbatot** belong to a month: all CSV rows sharing that `hyrmo`.
- **Week label (שבוע)**: derived as `clean(parsha_this) - clean(parsha_next)` across
  the whole file, so the last week of a month reaches into the next month's first
  parsha. `clean()` strips the `- מברכים` suffix. The very last row of the dataset has
  no next week, so its label is single and its weekday mincha is blank (no next-week data).
- **daf = 19:45** and **arvit2 = 20:30**: constants (confirmed constant all year in the XLS).
- **shacharit 2 = 8:30**, **shacharit 1 = 6:45** (or the `shacharit1` override): constants.
- **YomText (special events)** and **Rosh Chodesh**: no CSV source yet → rendered **empty**.
  (YomText = future manual column; Rosh Chodesh pending decision, time = fixed constant if kept.)

This unlocks any month the CSV covers (e.g. `--hyrmo 578612` = Elul 5786), which the XLS
could not supply.

### Year-boundary lookahead (last week of the dataset)

The final week in the CSV (last Shabbat before the next Rosh Hashana) has no "next week"
row, which would leave two gaps. Both are filled by looking ahead one Shabbat:

- **Week label**: `next_shabbat_label()` uses **pyluach with `israel=True`** to get the next
  Shabbat's parsha (or its holiday name when a holiday falls on it — e.g. `ראש השנה`).
  So the final label reads `נצבים-וילך - ראש השנה`.
  ⚠️ **Must use `israel=True`** — pyluach's default diaspora schedule diverges from the Modiin
  calendar by up to a week (verified: 41/43 match in Israel mode; the 2 diffs are spelling
  variants `שלח`/`שלח לך`, `פינחס`/`פנחס`, not schedule differences).
- **Weekday mincha**: when the CSV has no row for the "week after" (join key = Shabbat + 1 day),
  `weekday_mincha_for_week()` computes it on the fly with **ephem**, using the identical rule to
  `modiin_weeks.py` (`floor_to_5(earliest Sun–Thu sunset − 12)`). Reuses the calibrated helpers,
  so it equals the value next year's CSV would produce. Also self-heals any mid-year gap.

### Parsha column — name is derived from the DATE, not the PDF

The PDF's parsha text is unreliable on holiday Shabbatot and its wording drifts year to year
(e.g. the 5787 PDF gave `א` for 1 Tishrei). So `display_name()` in `yakir_calendar.py` derives the
displayed name from the date via pyluach, with this precedence:
1. **Chol HaMoed** (see below) → `שבת חוה"מ סוכות` / `שבת חוה"מ פסח`.
2. **Any other holiday on Shabbat** → the holiday name (remapped via `HOLIDAY_NAME_MAP`), e.g.
   1 Tishrei → `ראש השנה`, Shemini Atzeret → `שמחת תורה`, Sukkot 1st day → `סוכות`.
3. **Ordinary Shabbat** → the cleaned PDF parsha name.

This is not hard-coded to any one holiday — it fixes every holiday-on-Shabbat where the PDF text is
garbled or inconsistent.

### Shabbat Chol HaMoed (Sukkot / Pesach)

The Modiin PDF labels the Shabbat during Chol HaMoed with the Torah-reading name **`ויאמר משה`**
(the Ki Tisa Chol HaMoed reading), which is confusing in the parsha column. `chol_hamoed_shabbat()`
detects these via pyluach (Saturday + holiday `סוכות` Tishrei 16–21, or `פסח` Nissan 16–20) and:
- **Relabels the parsha** → `שבת חוה"מ סוכות` / `שבת חוה"מ פסח` (short form fits the column; also
  used in the adjacent week-labels).
- **Overrides shacharit-1** → Sukkot **6:15**, Pesach **6:30** (default is 6:45). The "bold when ≠ 6:45"
  rule then bolds these, correctly flagging them as exceptions.

Note: the Sukkot **1st day** on Shabbat (Tishrei 15) is NOT Chol HaMoed, so it shows the holiday name
`סוכות` (rule 2 above) — but it still gets שחרית 6:15 and no 17:15, via `sukkot_shabbat()` (§5).



Rendered below the weekday Shacharit as accent-colored lines (the month's `YomText`, `|`-separated).
Computed in `yakir_calendar.py` `month_events(heb_year, tishrei_month_num)` via pyluach
(**`israel=True`** — diaspora schedule is wrong for Modiin). Format: `שם - יום ד.מ`
(name - Hebrew weekday D.M), e.g. `יום כיפור - חמישי 2.10`.

**What is shown (this phase):** major holidays + Rosh Chodesh. **Fasts deferred** (kept in
`EVENT_FASTS`, unused; enable by merging into `EVENT_HOLIDAYS`).

**Holiday rules:**
- Single representative day (day 1) for multi-day chagim — סוכות, שמחת תורה, פסח show one day.
- **ראש השנה** is the ONLY holiday shown as a 2-day range (candle-lighting both days),
  e.g. `ראש השנה - שלישי-רביעי 23-24.9`.
- **שמיני עצרת → "שמחת תורה"** (Israel same day; `HOLIDAY_NAME_MAP`).
- **שביעי של פסח** injected separately (21 Nissan) — pyluach lumps all 7 Pesach days as "פסח",
  so the 7th yom tov is added explicitly.

**Rosh Chodesh rules** (the calendar is published before the first Shabbat, so it must carry the
NEXT month's start too):
- Show **both** the current month's RC and the next month's RC (same RC appears on two consecutive
  months — once as "next", once as "current" — intended).
- RC spans 1 or 2 days by the previous month's length (30-day month → 2-day RC).
- **Skip** the current month's RC if any of its days is **Shabbat** (visible in the Shabbat row) or
  if it coincides with **Rosh Hashana** (ראש השנה takes precedence).
- **Tishrei exception**: on the **Elul** page, the "next month's RC" (= 1 Tishrei) is shown as
  **"ראש השנה"**, and omitted if 1 Tishrei is Shabbat.

**Source:** automatic only (no manual override this phase). An override could later short-circuit
`month_events()` without changing the computation.

**Fasts** are included: `EVENT_FASTS` (צום גדליה, י׳ בטבת, תענית אסתר, י״ז בתמוז, ט׳ באב) is merged
with `EVENT_HOLIDAYS` in `month_events()` — same `שם - יום ד.מ` format, single day.

### Special-Shabbat tags (four parshiot)

`special_shabbat_tag()` uses pyluach `four_parshios()` to tag the parsha cell with שקלים / זכור / פרה /
החודש, rendered as `parsha - tag` (tag in a smaller font, same paragraph). Only the four parshiot
pyluach detects for free; שבת מברכים is not tagged (always the month's last week). Skipped on Chol
HaMoed Shabbatot (already relabeled).

### Conditional notes

- **Chanukah note** (Kislev only): a 2nd left-cell line under the arvit footnote —
  `בחנוכה תפילת ערבית ראשונה כ-20 דקות אחרי השקיעה` (`CHANUKAH_NOTE`). Detected by month **name**
  (`"כסלו" in chodesh`), not number.
- **School-vacation note** (`SCHOOL_NOTE`): under the Shacharit table (supports it — gap after, not
  before). Shown on Tishrei + Nissan always, plus any month with a Shabbat date in **21 Jun – 31 Aug**
  (→ Tammuz, Av, Elul).

  > **Leap-year rule (important):** Tishrei/Nissan are detected by month **NAME** (`"תשרי"`/`"ניסן"`
  > in `month["chodesh"]`), NOT by number. In a leap year Nissan is month **8** (7 = Adar II), so the
  > old numeric `heb_mo in (1,7)` put the note on אדר ב׳ and missed Nissan. Any month-specific note
  > must key on the name, never the Tishrei-based number.

### Label

The מנחה sub-table title is **`מנחה וערבית`** (was `מנחה, דף יומי וערבית`). If the mincha table is
ever reverted to per-row daf/arvit constants, restore the longer title too.

---

## 3. The Three Tools

### Tool 1: parse_modiin_pdf.py

**Input**: Modiin annual candles PDF (one page, RTL table)  
**Output**: `modiin_calendar.csv`

Parses text spans by x-coordinate from the PDF:

| Column x-range | Field |
|---|---|
| x ≤ 139 | motzash |
| x ≤ 184 | shkia |
| x ≤ 229 | candles |
| x ≤ 326 | parsha |
| x ≤ 374 | date_g (Gregorian DD/MM/YYYY) |
| x > 374 | date_h (Hebrew date string) |

Stops at the fast/miqveh section (text "צום", "כניסת הצום", "צאת הצום").

**Double-parsha normalization** (`normalize_parsha()`): Hebrew is RTL, so a combined
reading like נצבים-וילך (Nitzavim read first, then Vayelech) is extracted from the PDF
in **reversed** logical order with irregular whitespace — e.g. `וילך   -נצבים`. The
parser splits on the dash, strips, reverses to correct Torah reading order, and rejoins
with a single clean hyphen. All 6 double-parshiot in a year are affected. Verified 6/6:

| PDF extracts (wrong) | Normalized to |
|---|---|
| `וילך   -נצבים` | `נצבים-וילך` |
| `פקודי  -ויקהל` | `ויקהל-פקודי` |
| `מצורע  -תזריע` | `תזריע-מצורע` |
| `קדושים  - אחרי מות` | `אחרי מות-קדושים` |
| `בחוקותי  -בהר` | `בהר-בחוקותי` |
| `מסעי  -מטות` | `מטות-מסעי` |

> **Next year's PDF**: this normalization runs automatically — no manual fix needed.
> If a *new* double-parsha appears, verify it lands in correct reading order (the rule
> is generic: reverse the two dash-separated names). Single parshiot pass through unchanged.

Holiday detection keywords → codes:

| Hebrew text | Code |
|---|---|
| ראש השנה | RH |
| יום כיפור | YK |
| סוכות | SUK |
| שמחת תורה | SIM |
| שמיני עצרת | SMI |
| פסח / שביעי של פסח | PES |
| שבועות | SHV |

**Verification output**: Prints a diff table comparing official PDF times vs. a
simple `shkia−18min` / `shkia+42min` formula. This is FOR INFORMATION ONLY — the
official PDF times always win.

Usage:
```bash
python parse_modiin_pdf.py --pdf path/to/Modiin-2025-6-Candels.pdf
```

---

### Tool 2: populate_xls.py

**Input**: `modiin_calendar.csv` + existing XLS workbook  
**Output**: Updated XLS with Detail sheet populated

Writes these columns in the Detail sheet:

| Column | Description | Format stored |
|---|---|---|
| AJ (col 36) | Candles (official, from PDF) | `datetime.time` in 12H (subtract 12h) |
| AK (col 37) | Motzash (official, from PDF) | `datetime.time` in 12H (subtract 12h) |
| AL (col 38) | Shkia | `datetime.datetime` (date + time) |
| AM (col 39) | Date | `datetime.datetime` |
| AO (col 41) | Parsha Input | string |
| AP (col 42) | Next Parsha | string (from next CSV row) |
| AQ (col 43) | Holiday code | string (RH, YK, PES, etc.) |
| AR (col 44) | Gregorian year | int |
| AS (col 45) | Summer Y/N | (legacy XLS) "Y" if candles ≥ 17:30. Live pipeline uses candles > 18:00 — see §Summer Detection Threshold |

**12H storage quirk**: The XLS stores AJ/AK as 12H times (so 18:03 is stored as
`time(6, 3)`). This is how the XLS was designed — the formulas add 12h back when needed.

**Verification diff**: Before writing, the tool prints a comparison of existing XLS
values vs. CSV values. Differences are flagged. The CSV values are always written
(they are authoritative).

Usage:
```bash
python populate_xls.py --csv modiin_calendar.csv --xls path/to/workbook.xlsx
python populate_xls.py --dry-run  # preview without saving
```

---

### Tool 3: yakir_calendar.py

**Input**: XLS `ResultForWordMerge` sheet  
**Output**: `yakir_calendar_output.docx` (and optionally `.pdf`)

Reads 15 rows (months) from the sheet. Each row has up to 5 Shabbat blocks
(prefixes a–e) plus weekday data and Rosh Chodesh info.

Usage:
```bash
python yakir_calendar.py --year 5787           # WHOLE year (12 or 13 months — leap-year auto)
python yakir_calendar.py --hyrmo 578701 578702 # one/few months by year+month key (Tishrei=01)
python yakir_calendar.py                       # all rows in the CSV (incl. partial trailing month)
python yakir_calendar.py --pdf                 # also convert to PDF
python yakir_calendar.py --out my_output.docx
```
Selection precedence: `--year` > `--hyrmo` > `--months` (0-based, fragile) > all. `--year` figures out
12-vs-13 months from the CSV and warns if the year isn't fully covered.

---

## 4. Critical Rule: Candles & Motzash Are Sacred

> **CANDLES AND MOTZASH VALUES MUST ALWAYS COME FROM THE OFFICIAL MODIIN CITY PDF.**  
> Python code must NEVER calculate, override, or substitute these values.

The reason: the Modiin municipality publishes official candle-lighting and Havdalah
times that account for local factors (hills, microclimate, rabbinic decisions). The
simple formula `shkia − 18min` is only an approximation used for verification.

In practice, Modiin uses approximately `shkia − 20min` for candles and `shkia + 30min`
for Motzash — but even this is not precise. The PDF is always authoritative.

**Verification is allowed**: `parse_modiin_pdf.py` prints a diff table showing how
much the official times deviate from a calculated formula. This helps sanity-check
that the PDF was parsed correctly. It does not change any values.

---

## 5. Prayer Time Rules

### Shabbat Prayer Times (now sourced from modiin_weeks.csv)

The Shabbat block is computed in `modiin_weeks.py` and rendered by `yakir_calendar.py`
from `modiin_weeks.csv` (keyed by `weekend_date`). The XLS is validation-only. Each
computed time has a `_calc` column plus a `_adj` manual-override column (adj wins).

| Time (column) | Rule | Source |
|---|---|---|
| הדלקת נרות (`candles`) | FROM PDF ONLY — never calculated | Modiin PDF |
| מוצאי שבת (`motzash`) | FROM PDF ONLY — never calculated | Modiin PDF |
| מנחה שבת ktana (`shabbat_mincha_ktana_calc`) | `floor_to_5(candles)` | computed + `_adj` |
| מנחה וקבלת שבת (erev shabbat) | `mincha_ktana + 10 min` — **derived at render, not stored**; tracks any ktana override | derived (generator) |
| סוף זמן קריאת שמע (`sof_zman_shma_calc`) | GRA: `sunrise + (sunset−sunrise)/4` via ephem (MAE 0.36 min vs OU) | computed + `_adj` |
| שחרית שבת 1 (`shacharit1`) | blank → default **6:45**; fill to override (e.g. Rosh Chodesh 6:30) | manual (default in code) |
| תאריך עברי (`taarich`) | Hebrew day + month, e.g. `ה׳ תשרי` | pyluach |

**Constants (in code, NOT CSV columns):**
- מנחה גדולה **13:30** — always shown
- Sharbat mincha **17:15** — summer early mincha (organized by Yair Sharbat), shown when `sharbat_mincha = yes` — **except** on Sukkot / שבת חוה"מ / שמחת תורה (suppressed, see below)
- שחרית שבת 2 **8:30** — second minyan

**Festival Shabbat overrides** (date-based, via pyluach — correct every year, no manual dates):
| Shabbat (by date) | Effect | Helper |
|---|---|---|
| Sukkot, Tishrei 15-21 (1st day OR chol hamoed) | שחרית 1 → **6:15** (Kohelet is long); **no 17:15** Sharbat mincha | `sukkot_shabbat()` |
| שבת חול המועד פסח (Nissan 16-20) | שחרית 1 → **6:30** | `chol_hamoed_shabbat()` |
| שמחת תורה (Tishrei 22) **or** Pesach 1st day (Nissan 15) on Shabbat | שחרית 1, שחרית 2, **and** the מנחה שבת string all left **BLANK** (special festival timetable, filled manually); no 17:15 | `yomtov_no_minyan_shabbat()` |

The `_adj` columns still handle any other one-off holiday fix. Other holidays (Yom Kippur, etc.)
use the congregation's dedicated timetables and are not rendered on this monthly sheet.

**Mincha erev shabbat is derived, not stored.** It equals `mincha_ktana + 10 min`, computed
by `yakir_calendar.py` at render time from the *resolved* ktana (adj if set, else calc). This
matches the old calendar exactly (erev = ktana + 10, verified 50/50 weeks) and means a manual
`shabbat_mincha_ktana_adj` override automatically shifts erev too — one source of truth.

### Mincha Shabbat — The Summer Rule (display string)

The מנחה שבת cell is assembled by `format_shabbat_mincha()` in `yakir_calendar.py`:

**Winter** (`sharbat_mincha = no`): `"13:30  HH:MM"` — gedola + mincha ktana
**Summer** (`sharbat_mincha = yes`): `"13:30  17:15  HH:MM"` — gedola + Sharbat 17:15 + mincha ktana

**Detection**: the `sharbat_mincha` column (yes when candles **> 18:00**). Replaces the old
XLS `earlyMin`-based detection.

### Weekday Prayer Times

Weekday times in the `תפילות יום חול` section come from the `wk*` fields:

| Field | Description |
|---|---|
| `wkMin` | Weekday mincha time |
| `wkArv1` | Weekday arvit 1 (right after mincha) |
| `daf` | Daf Yomi shiur time |
| `wkArv2` | Weekday arvit 2 (later evening) |
| `wkLbl` | Week label ("פרשה1 - פרשה2" format) |

The weekday section shows:
1. Last Shabbat of previous month (`lastmo-*` fields) — for the first week of this month
2. Each Shabbat's weekday block for subsequent weeks

### Shacharit
- **Sunday–Thursday**: 6:45
- **Friday**: 8:30 (second minyan) / 6:15 (first minyan)
- **Rosh Chodesh / Hanukkah / special days**: 8:30 (second minyan added)

---

## 6. XLS Architecture

### Workbook: `yakircal master v61 Added shkia+14 times starting 5785.xlsx`

| Sheet | Purpose |
|---|---|
| `Detail` | Raw input layer. One row per Shabbat. Columns AJ-AS = raw inputs; all other columns = formula-calculated |
| `ResultForWordMerge` | Pre-calculated output layer. 15 rows (months) × up to 5 Shabbatot per row. This is what the Word generator reads. |
| `sunset` | Daily sunrise/sunset data with full mincha calculation logic |
| `Key` | Documents all Detail column meanings and calculation rules |
| `shma_time` | Sof Zman Shma lookup table |
| `yom` / `holidays` | Holiday detection and handling tables |
| `RoshChodesh` | Rosh Chodesh date lookup |

### ResultForWordMerge Column Structure

Each month row has 123 columns:

- **Shabbat blocks**: prefixes `a-` through `e-` (23 columns each = 5 × 23 = 115 cols)
- **lastmo-*** (5 cols): weekday data from last Shabbat of previous month
- **YomText**: pipe-separated special events string for this month
- **ThisRC**: Rosh Chodesh info for this month
- **NextRC**: Rosh Chodesh info for next month

Per-Shabbat fields (prefix `a-` through `e-`):

| Field | Type | Description |
|---|---|---|
| `hyrmo` | int (YYYYMM) | Year-month key |
| `wk` | int | Shabbat number within month (1-5) |
| `mo` | int | Hebrew month number (1=Tishri) |
| `hyr` | int | Hebrew year number |
| `shana` | str | Hebrew year in letters (תשפ״ו) |
| `chodesh` | str | Hebrew month name |
| `motzsh` | time | Motzash time — FROM PDF ONLY |
| `min1` | time | First (early) mincha time |
| `min2` | str | Pre-formatted mincha display: "13:30  HH:MM" |
| `shach2` | str | Shacharit 2 time (8:30) |
| `shma` | time | Sof Zman Shma |
| `shach1` | str | Shacharit 1 time (6:45 / 6:30 / ---) |
| `friMin` | time | Friday mincha (entry time for Shabbat) |
| `earlyMin` | time or "--" | Summer early mincha (17:15) or "--" |
| `candles` | time | Candle lighting — FROM PDF ONLY |
| `parsha` | str | Parasha name |
| `date` | datetime | Gregorian date of Shabbat |
| `taarich` | str | Hebrew date string (e.g. ה׳ תשרי) |
| `wkArv1` | time | Weekday arvit 1 (for following week) |
| `daf` | time | Daf Yomi time |
| `wkArv2` | time | Weekday arvit 2 |
| `wkMin` | time | Weekday mincha |
| `wkLbl` | str | Week label "parsha1 - parsha2" |

---

## 7. Output Document Layout

### Page: A4 Landscape (841 × 595 pt)

```
┌─ Margin 1cm ──────────────────────────────────────── Margin 1cm ─┐
│                                                                   │
│  [logo 3.5cm]    לו״ז חודש [CHODESH] [SHANA] 28pt bold   בס״ד   │
│                                                                   │
│                      זמני השבת  22pt bold                        │
│  ┌─────┬─────────┬──────┬───────┬───────┬────────┬───────┬...┐  │
│  │מוצאי│ מנחה שבת│סוף זמן│שחרית 2│שחרית 1│מנחה+קבל│הדלקת│...│  │
│  │שבת  │         │קריאת  │       │       │ שבת    │נרות │...│  │
│  │וערב │         │ שמע   │       │       │        │     │...│  │
│  ├─────┼─────────┼──────┼───────┼───────┼────────┼───────┼...┤  │
│  │19:03│13:30 17:15│09:31│  8:30 │  6:45 │ 18:10  │18:13 │...│  │
│  │     │   18:00 │      │       │       │        │      │...│  │
│  │ ... │  ...    │ ...  │  ...  │  ...  │  ...   │  ... │...│  │
│  └─────┴─────────┴──────┴───────┴───────┴────────┴───────┴...┘  │
│                                                                   │
│                    תפילות יום חול  20pt bold                     │
│  ┌──────────────────────┬──────────────────────────────────────┐ │
│  │       שחרית           │    מנחה, דף יומי וערבית             │ │
│  │ ראש חודש: ...         │  שבוע         מנחה  דף   ערבית2    │ │
│  │ יום א׳–ה׳: 6:45       │  פרשה1-פרשה2  18:05 19:45  20:30   │ │
│  │ יום ו׳: 8:30 / 6:15   │  ...                                │ │
│  └──────────────────────┴──────────────────────────────────────┘ │
│                                                                   │
│  [YomText events if any — one line per event]                    │
│                                                                   │
│  * מניין ראשון לתפילת ערבית מתקיים לאחר סיום תפילת מנחה (ולאחר השקיעה) │
│  [org info lines]                                                 │
└───────────────────────────────────────────────────────────────────┘
```

### Fonts

Names are pinned unconditionally: **Arial** (`F_ARIAL`) and **Calibri** (`F_CALIBRI`) — see the
Font-name rule below. Sizes (Hebrew is complex-script → `_run` sets both `w:sz` AND `w:szCs`, else
Word falls back to 11pt):

| Element | Font | Size | Style |
|---|---|---|---|
| Title (לו״ז חודש...) | Arial | 28pt | Bold |
| בס״ד | Arial | 16pt | Bold |
| Section headers (זמני השבת / תפילות יום חול) | Arial | 22 / 20pt | Bold |
| Sub-table titles (מנחה וערבית / שחרית) | Arial | 15pt | Bold |
| Column headers & data cells | Calibri | 16pt | Normal |
| Events (holidays/RC) | Calibri | 12pt | Bold, accent color |
| Footnote / notes | Arial | 11-12pt | Normal |
| daf/arvit line | Calibri | 18pt | Bold |
| Org info | Arial | 9pt | Mixed |

### Shabbat Table — 10 Columns (RTL order, left-to-right visually)

Widths are relative units in `SHAB_COLS`, rescaled to `BODY_W` (`w * BODY_W / _total`); with
`_total == BODY_W == 27.7` a unit ≈ 1 cm. Current values:

| Col | Header | Width |
|---|---|---|
| 0 | מוצאי שבת / וערבית | 2.7 |
| 1 | מנחה שבת | 5.0 (holds `13:30  17:15  HH:MM`) |
| 2 | סוף זמן קריאת שמע | 2.1 |
| 3 | שחרית שבת (8:30) | 2.1 |
| 4 | שחרית שבת (6:45) | 2.1 |
| 5 | מנחה וקבלת שבת | 2.2 |
| 6 | הדלקת נרות | **2.3** (only ever a time → the table's idle slack) |
| 7 | פרשת השבוע | 4.0 (must fit `מטות-מסעי` + `א' דראש השנה`) |
| 8 | תאריך לועזי | 2.1 |
| 9 | תאריך עברי | **3.1** (must fit leap-year `י״ג אדר א׳` on one line) |

> **Column-width rule (why Tishrei once spilled to 2 pages):** `SHAB_COLS` is ONE fixed layout for
> ALL months, so each column must be sized for its **year-wide worst case**, not the month you're
> viewing. תאריך עברי was 2.3 (too narrow for leap-year Adar dates) → the date wrapped to 2 lines →
> `exact=False` rows grew → a 5-Shabbat month overflowed. Fixed by giving it 3.1, taken from הדלקת
> נרות (which only shows a time). **Never** widen a column by shrinking פרשת השבוע — it's the
> tightest text column. Keep the widths summing to `BODY_W` so no column shifts.

> **Font-name rule (cross-machine consistency):** a `.docx` does NOT embed fonts — it names them, and
> Word renders with whatever that name resolves to on the machine that OPENS the file. So `F_ARIAL`/
> `F_CALIBRI` are pinned **unconditionally** ("Arial"/"Calibri") — never branch the name on the build
> host. (Old bug: a WSL-vs-Windows check named "DejaVu Sans" off-WSL, whose wide Hebrew glyphs
> changed line wrapping and pagination.) Arial → Liberation Sans on bare Linux (metric-compatible).

---

## 8. Technical Appendix for AI

This section is written for an AI agent that may need to modify, debug, or extend this
codebase in the future.

### Repository Structure

```
Yakir-Calendar-creator/          ← all files in this folder
  parse_modiin_pdf.py            ← Tool 1: PDF → CSV
  populate_xls.py                ← Tool 2: CSV → XLS Detail sheet
  yakir_calendar.py              ← Tool 3: XLS ResultForWordMerge → Word
  yakir_logo.png                 ← Logo extracted from original PDF
  RULES.md                       ← This file

../                              ← Yakir-Calendar-experimental/
  modiin_calendar.csv            ← Output of Tool 1 (59 rows)
  yakircal master v61 ....xlsx   ← The master XLS workbook
  yakir_logo.png                 ← Original logo file
  Yakir-calendar-2025-6/         ← Source PDFs
    Modiin-2025-6-Candels.pdf    ← City candles PDF (input to Tool 1)
    Yakir 5786 all months.pdf    ← Example output PDF (12-page reference)
```

### XLS File Path

```
yakircal master v61 Added shkia+14 times starting 5785.xlsx
```

Default path in all tools: `../yakircal master v61 Added shkia+14 times starting 5785.xlsx`
(relative to the Yakir-Calendar-creator folder).

### Key openpyxl Notes

- Open with `data_only=True` to read formula-calculated values (not formula text)
- `datetime.time` objects → use `.strftime('%H:%M')` for display
- `datetime.datetime` objects from date cells → use `.strftime('%d/%m/%Y')`
- AJ/AK are stored as 12H time — to convert to display: add 12 to the hour
  ```python
  def to_24h(t):  # t is datetime.time stored in XLS (12H format)
      return datetime.time(t.hour + 12, t.minute)
  ```

### Key python-docx Notes

- Hebrew RTL requires marking both paragraphs (`w:bidi`) and runs (`w:rtl`)
- Font must be set on both `rFonts.ascii/hAnsi` and `rFonts.cs` (complex script)
- Table cell text must be set via `cell.paragraphs[0]`, not `cell.text = ...`
  (the latter removes existing paragraph format)
- Column widths must be set per-row (not per-column object) via `w:tcW`
- For A4 landscape: `page_width=Cm(29.7)`, `page_height=Cm(21.0)`

### Adding a New Year

1. Get the new Modiin candles PDF for the new year
2. Run `parse_modiin_pdf.py --pdf new_pdf.pdf --out new_data.csv`
3. Run `populate_xls.py --csv new_data.csv --xls path/to/workbook.xlsx --year NNNN`
4. Verify the XLS formulas recalculated correctly (open in Excel)
5. Run `yakir_calendar.py --out new_year_output.docx`

### Common Issues

| Symptom | Cause | Fix |
|---|---|---|
| Hebrew text appears as boxes | Font doesn't have Hebrew glyphs | Install Windows fonts or use DejaVu Sans |
| Wrong candles time in output | CSV not updated / wrong year | Re-run parse_modiin_pdf.py with correct PDF |
| Table overflows page | Too many columns or large font | Reduce font size or column widths in yakir_calendar.py |
| Motzash shows wrong time | XLS formula needs recalculation | Open in Excel, press Ctrl+Alt+F9 to force recalc |
| 12H / 24H confusion | AJ/AK are always stored as 12H in XLS | Use `to_12h()` when writing, add 12h when reading for display |
| LibreOffice PDF broken RTL | Known LibreOffice RTL issue | Use Word on Windows for PDF via PowerShell COM |
| `data_only=True` returns None | XLS not been saved after formula changes | Open in Excel, save, then run again |

### Candles/Motzash Verification Formula

The official Modiin formula is approximately:
- Candles ≈ shkia − 20 minutes (not 18)
- Motzash ≈ shkia + 30 minutes (not 42)

The simple verification formula (`shkia−18` / `shkia+42`) is used only to detect
obvious parsing errors. A typical diff of ±2 minutes is normal and expected.

### Summer Detection Threshold

> **Current rule (live pipeline)**: `modiin_weeks.py` sets `sharbat_mincha = yes` when
> candles are **later than 18:00** (`SUMMER_CANDLES_THRESHOLD = (18, 0)`, strict `>`).
> The paragraph below describes the **legacy** `populate_xls.py` behavior (XLS path,
> retired) and is kept for reference only.

`is_summer()` (legacy, populate_xls.py) returned "Y" if `candles >= 17:30`.

This aligned with the old XLS rule: when candles are late enough that a fixed early
mincha at 17:15 is appropriate. The live pipeline supersedes this with the 18:00 rule above.

### YomText Format

The `YomText` field in ResultForWordMerge uses pipe `|` as separator.
Each segment is a free-text description of a special event for that month.
Example:
```
ראש השנה (+ערב ר״ה שח830): 3/10-4/10|צום גדליה: 6/10|שבת יום כפור: 12/10
```

Display: split on `|`, render each segment on its own line, right-aligned.

---

*Last updated: 2026-07*
