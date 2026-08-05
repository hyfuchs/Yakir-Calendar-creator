# SESSION NOTES — Yakir Calendar Creator
# Last updated: 2026-08-04

## All files location
`/mnt/c/yfuchs/OneDrive/Documents/Yakir-Calendar-experimental/Yakir-Calendar-creator/`
Per-year data + outputs now live in a year folder, e.g. `../Yakir-calendar-2026-7/`.

---

## 2026-08-05 SESSION (part 12) — holiday parsha names fixed at PARSE time

### Problem
5787 CSV first row (1 Tishrei) showed `parsha=א`, `holiday=` (empty). The municipality PDF is
inconsistent year to year; for 5787 it wrote just `א` for Rosh Hashana, so the text-only parser
mislabeled it and `detect_holiday` found no keyword. The render layer (`display_name`) already
corrected the *printed* calendar, but the *CSV* was wrong.

### Fix — `parse_modiin_pdf.py` only
- Added a **guarded, optional** pyluach import (WARN + fall back to PDF text if missing — never
  hard-fails, unlike make_skeleton_csv.py which requires it).
- New helper `holiday_name_for_date(date_g)`: returns the pyluach Israel holiday name ONLY when the
  date both (a) has no weekly parsha (`getparsha_string is None`) AND (b) `holiday()` is truthy.
  This AND-predicate is the key safety: it isolates the displacing yom-tovim (RH/YK/Sukkot/Shemini
  Atzeret/Pesach/Shavuot) while leaving alone ordinary parshiot AND minor-holiday Shabbatot
  (Chanukah, Tu BiShvat) that keep a normal parsha.
- Wired into `parse_pdf`: on a holiday date use the pyluach name (raw — NO `normalize_parsha`, NO
  `HOLIDAY_NAME_MAP`; render layer owns display remaps); otherwise keep the existing
  `normalize_parsha(_clean_quotes(parsha_raw))` path unchanged. `detect_holiday(parsha)` then yields
  the code for free.

### Why low risk (verified)
- Rendered calendar is byte-identical whether CSV holds old PDF text or new pyluach name (render
  transforms sit on top) — so this is cosmetic to the CSV, cannot change what prints.
- Did NOT blanket-replace parsha: pyluach spells some ordinary parshiot differently (שלח לך→שלח,
  פנחס→פינחס) — those are the congregation's authoritative PDF spellings and are preserved.
- Verified 5787: 12/09→ראש השנה/RH, 21/09→יום כיפור/YK, 26/09→סוכות/SUK, 03/10→שמיני עצרת/SMI,
  24/04→פסח/PES; שלח לך, פנחס, וישב (Chanukah), בשלח (Tu BiShvat) all UNCHANGED. 5786 regen: 59 rows,
  ordinary parshiot intact. Full pipeline (weeks→render Tishrei) still builds; ±2min gate passes.
- Render-time `display_name()` fix kept as defense-in-depth.



### `--year` flag on yakir_calendar.py (whole-year generation)
- Was: type all hyrmo keys (`--hyrmo 578701 … 578713`) and you had to KNOW it's a leap year (13 keys).
- Now: `--year 5787` selects every month of that Hebrew year from the CSV, sorted Tishrei→Elul.
  Leap years handled automatically (12 or 13 months — whatever exists). Precedence: `--year` >
  `--hyrmo` > `--months` > all. Warns if the CSV doesn't cover a full year (prints month count found).
- `--hyrmo` retained for rendering one/few months (quick review). HOW_TO_RUN updated: `--year` is the
  headline; the old 13-key list is gone.

### Font-name build-host bug (🤖 caused wrong pagination off-WSL)
- `.docx` does NOT embed fonts — it only NAMES them; Word renders with whatever that name resolves to
  on the machine that OPENS the file (always Windows Word here).
- BUG: `HAS_WIN = os.path.exists("/mnt/c/Windows/Fonts")` made the font name DEPEND ON THE BUILD HOST.
  WSL builds named "Arial" (good); native-Windows builds named "DejaVu Sans" (wide/absent Hebrew
  glyphs) → the long week label wrapped → month spilled to page 2. Same input, different output by
  machine.
- FIX: pin `F_ARIAL="Arial"`, `F_CALIBRI="Calibri"` UNCONDITIONALLY (removed HAS_WIN). Arial is on
  Windows+Mac; on bare Linux LibreOffice substitutes metric-compatible Liberation Sans → line widths
  and pagination stay identical everywhere. **Rule: never branch a font name on the build host.**

### Tishrei 2-page overflow → widen תאריך עברי column
- Root cause (user-diagnosed, confirmed by width math): the תאריך עברי column (2.3 units ≈ 2.05 cm,
  ~65pt) was too narrow for leap-year Adar dates (`י״ג אדר א׳` needs ~80pt) → the date WRAPPED to 2
  lines. Those data rows are `exact=False` (height is a floor), so each wrapped row grew; in 5-Shabbat
  Tishrei the last 3 dates wrapped → ~1 cm growth → page 2. Not a "heavy month" problem.
- FIX: in `SHAB_COLS`, swap two width numbers only — תאריך עברי 2.3→**3.1**, הדלקת נרות 3.1→**2.3**
  (candles only ever shows a time like 18:42, so it had +36pt idle slack). Total unchanged (27.7) →
  NO other column shifts; order untouched. פרשת השבוע (4.0) deliberately NOT touched — it must fit
  long double-parshiot (`מטות-מסעי`) + `א' דראש השנה`; it's the tightest text column, off-limits as a
  donor. Verified: 5787 → exactly 13 pages, no wrap in any month incl. Adar I/II.
- **Rule: SHAB_COLS is one fixed layout for ALL months — size each column for its YEAR-WIDE worst
  case, not the month you're looking at. Idle slack lives in הדלקת נרות (always just a time).**

### Festival prayer-time fixes (Tishrei + Pesach)
Two new date-based helpers (pyluach, so correct every year):
- `sukkot_shabbat(d)` — Shabbat in Tishrei 15-21 (Sukkot 1st day OR chol hamoed).
- `yomtov_no_minyan_shabbat(d)` — Shemini Atzeret/Simchat Torah (Tishrei 22) OR Pesach 1st day (Nissan 15).
Applied in build_page data-row loop:
1. **Sukkot shach1 = 6:15** (Kohelet is long) — now covers the 1st day too, not only chol hamoed
   (which `chol_hamoed_shabbat` already set: Sukkot 6:15 / Pesach 6:30). 5787: 26/09 סוכות → 6:15.
2. **שמחת תורה / Pesach-1st-on-Shabbat → BLANK** shacharit1, shacharit2, AND the Shabbat mincha
   string (special festival timetable, filled manually). 5787: 03/10 שמחת תורה → all three empty.
   (Original idea of a "?" placeholder was dropped by the user in favor of blank.)
3. **No 17:15 (Sharbat)** on Sukkot / שבת חוה"מ / שמחת תורה — `format_shabbat_mincha(blk,
   suppress_sharbat=True)`. Pesach was explicitly left UNAFFECTED (user choice). Summer weeks keep 17:15.

### Holiday name derived from date, not PDF (parsha column)
- The PDF's parsha text is unreliable on holiday Shabbatot and drifts yearly (it gave `א` for 1 Tishrei).
- FIX: `display_name()` now, for any Shabbat that pyluach marks a holiday, uses the (remapped) holiday
  name instead of the PDF text: 1 Tishrei → `ראש השנה`, Shemini Atzeret → `שמחת תורה` (via
  HOLIDAY_NAME_MAP). Chol HaMoed still → `שבת חוה"מ סוכות`/`שבת חוה"מ פסח`. Ordinary Shabbatot keep the
  PDF parsha. Not hard-coded to any one holiday.

### School-vacation note — leap-year bug (🤖)
- `show_school = heb_mo in (1, 7)` assumed Nissan = month 7. In a LEAP year Nissan = 8 (7 = Adar II),
  so the note landed on אדר ב׳ and missed Nissan.
- FIX: detect Tishrei/Nissan by NAME (`"תשרי"`/`"ניסן"` in month["chodesh"]), not by number. Also
  is_kislev switched to name-based (`"כסלו"`). Verified 5787: note now on תשרי/ניסן/סיון/תמוז/אב; Adar II clean.

### Windows batch file — `../Yakir-calendar-2026-7/build_yakir_5787.bat` (NEW)
- One-click full pipeline from the PDF: parse → weeks → full-year calendar (+PDF), then a separate
  Tishrei-only build (`--hyrmo 578701`, writes `yakir_5787_tishrei.docx`). Settings block at top
  (PDF/YEAR/THEME) for reuse; errorlevel checks; needs Word for the .pdf step. Runs the tools bare
  (cwd = year folder → inputs/outputs default there).

---

## 2026-08-04 SESSION (part 10) — new PDF format, content-based parser, skeleton fallback, 5787

### תשפ״ז PDF format changed → content-based parser
- New `Modiin-2026-7-Candels.pdf` broke the old coordinate parser (0 rows): different x-positions,
  and candles+parsha+dates MERGED into one span; curly quotes; spans 2 Hebrew years.
- **Rewrote `parse_modiin_pdf.py` to be CONTENT-based** (not coordinate-based): a data row = any row with
  a `DD/MM/YYYY` token; times assigned by CHRONOLOGY (candles<shkia<motzash); parsha/heb-date = leftover
  text (heb-date identified by containing a month name / "תשפ"). Handles both old & new layouts.
  Added `--debug`. Reuses `normalize_parsha` (double-parshiot still reversed in source) + `detect_holiday`.
- **Regression**: re-parsing last year's PDF reproduces the committed CSV EXACTLY (59 rows, 0 diffs).
- New PDF: 63 rows, chronology 0 bad, double-parshiot correct, ±2min gate passes.
- Design note: OCR rejected (PDF has real text layer); no LLM (deterministic regex on date/time anchors
  is durable to yearly layout drift). `holiday`/`date_h` CSV columns are NOT read downstream (pyluach
  supplies those) — only `date_g` + times matter.

### Skeleton fallback for image-only years — `make_skeleton_csv.py` (NEW)
- When only an image scan exists (no digital PDF): `make_skeleton_csv.py --year N` pyluach-generates all
  Saturdays (dates/parsha/holiday), leaves candles/shkia/motzash BLANK for typing, adds ephem `*_hint`
  columns to check typing against. Saturdays only (all the pipeline consumes).
- Verified: 5787 skeleton = 57 Sats, all dates match the real 2026-7 CSV; 7 parsha diffs all harmless
  (holiday/Chol-HaMoed/spelling — generator recomputes via pyluach). Fill→modiin_weeks→yakir_calendar works.
- ⚠️ double-parsha: pyluach already gives correct order — do NOT run normalize_parsha on it (that reverses).

### 5787 IS A LEAP YEAR (13 months)
- Has אדר א׳ (578706) + אדר ב׳ (578707). Full-year command needs **13** hyrmo keys 578701–578713.
- Generated `../Yakir-calendar-2026-7/yakir_5787.docx` + `.pdf` (13 months). PENDING user review.

### Layout note
- CSVs moved into the project dir earlier; per-year products go in the year folder (2026-7).
- Default paths in scripts now project-dir-relative (SCRIPT_DIR), not `../`.

### NEXT
- User to review yakir_5787.pdf (esp. the two Adar months, 5-shabbat months).
- User may share scanned images of past years (image only) — test skeleton path, not parser.
- All-year A3 table (separate future request).

---

## 2026-08-03 SESSION (part 8) — special-Shabbat tags, fasts, conditional notes, label

### Added (all in yakir_calendar.py, no CSV change)
1. **Special-Shabbat tags** — `special_shabbat_tag()` via pyluach `four_parshios()`: שקלים/זכור/פרה/החודש
   appended to the parsha cell as `parsha - tag` (tag smaller font, same paragraph). Skipped on Chol HaMoed.
   NOT מברכים (last week). Verified: שקלים→שבט 14/02, זכור/פרה/החודש→Adar.
2. **Fasts in events** — `EVENT_FASTS` merged into `month_events()` (`show = EVENT_HOLIDAYS | EVENT_FASTS`).
3. **Chanukah note** — Kislev only (heb_month_num==3, NOT 9 — hyrmo uses Tishrei=1). 2nd left-cell footnote line.
4. **School-vacation note** — `SCHOOL_NOTE` under Shacharit table (gap AFTER, snug to table). Tishrei+Nissan
   + summer months with a Shabbat in 21 Jun–31 Aug → Tishrei, Nissan, Tammuz, Av, Elul (5 months).
5. **Label rename** — `מנחה, דף יומי וערבית` → `מנחה וערבית` (~line 891; revert with mincha-table format).

### Gotcha fixed
- hyrmo month number is **Tishrei-based** (Kislev=3, not 9). First Chanukah-note attempt used 9 → never fired.

### Verified
- Adar tags, Shevat שקלים, Kislev Chanukah note (present; absent Cheshvan), school note (Tishrei/Nissan yes,
  Cheshvan no), Tishrei צום גדליה in events. All ✓.
- ⚠️ Page-fit: content added is small (1 tag line / 1 note line) — verify visually on full-year PDF.

### NEXT (deferred)
- שבת הגדול/שובה/נחמו (need manual date logic) — skipped per user.
- Events manual override; mincha-table revert to per-row constants (kept available).
- Folder cleanup: ~30 draft*.docx test files can be deleted, keep yakir_5786.docx/.pdf.

---

## 2026-08-03 SESSION (part 7) — Shabbat Chol HaMoed + RTL alignment bug + PDF + full-year

### Shabbat Chol HaMoed (Sukkot/Pesach)
- PDF labels these Shabbatot `ויאמר משה` (confusing). `chol_hamoed_shabbat(greg_date)` detects via
  pyluach (Sat + holiday סוכות Tishrei 16-21, or פסח Nissan 16-20) and returns name + shach1.
- Parsha relabel → `שבת חוה"מ סוכות` / `שבת חוה"מ פסח` (short form; applied to row parsha AND week labels
  via a `display_name()` helper in load_months_from_csv).
- Shacharit-1 override → Sukkot 6:15, Pesach 6:30 (in build_page, after CSV shacharit1 lookup). Auto-bolded.
- Verified: Tishrei 11/10 → שבת חוה"מ סוכות / 6:15; Nissan 04/04 → שבת חוה"מ פסח / 6:30. Cheshvan regression clean.

### CRITICAL BUG FIXED — RTL right-alignment (holidays rendered left)
- Root cause: an explicit `w:jc="right"` on an RTL (bidi) paragraph renders LEFT in Word. Confirmed by
  diffing a user-hand-fixed file: theirs had NO jc, just bidi → rendered right.
- Fix: events use `_cell_write(..., align=None)` → no jc emitted → bidi's natural right alignment wins.
  Also hardened `_bidi_para` to insert `<w:bidi>` in schema-correct order (before spacing/ind/jc).
- LESSON: for RTL cells, do NOT set jc=right; rely on bidi. Events render as a borderless bidiVisual
  single-column table in the שחרית (right) cell.

### PDF export FIXED
- `--pdf` flag now works. `to_pdf()` resolves the path to absolute first, converts /mnt/<drive>/ → Windows
  path correctly, drives Word via PowerShell COM ($w.Visible=$false). LibreOffice not installed; Word COM works.
- Command: `python3 yakir_calendar.py --theme purple --pdf --out X.docx` → produces X.docx + X.pdf.

### FULL YEAR
- CSV has 13 month-groups: אלול תשפ"ה (1 shabbat, prev-year tail) + Tishrei..Elul 5786.
- 12 months of 5786 only: `--hyrmo 578601 578602 578603 578604 578605 578606 578607 578608 578609 578610 578611 578612`
- All months / no filter → all 13.

### Layout (recap from parts 5-6, now stable)
- Events in שחרית cell; footnote + daf/arvit constants in מנחה cell (side-by-side, no overlap).
- Blank spacer lines before org footer: 2 for 4-shabbat, 1 for 5-shabbat month (10pt each, hand-deletable).
- All months fit one page (4- and 5-shabbat verified).

### NEXT (deferred)
- Special Shabbatot tags (שקלים, זכור, מברכים…) small-font next to parsha — still pending.
- Fasts in events section; events manual override.

---

## 2026-08-03 SESSION (part 6) — events section (holidays + Rosh Chodesh)

### What was built (yakir_calendar.py)
- `month_events(heb_year, tishrei_month_num)` — computes the month's event lines via pyluach
  (`israel=True`), returns `['שם - יום ד.מ', ...]`. Wired into `load_months_from_csv` → `YomText`.
- **Placement**: events render INSIDE the שחרית (right) cell of the weekday 2-col table, below the
  short shacharit table. Right-justified (jc=right + bidi=1), accent-colored, bold.
  The footnote (`* מניין ראשון...`) + daf/arvit constants line live in the LEFT cell (with the מנחה
  content). Left & right cells sit side-by-side → no overlap. Only org-info is full-width at the bottom.
- Helpers: `HEB_DOW`, `_heb_dow`, `_greg`, `_pyluach_month` (Tishrei→pyluach month), `_rosh_chodesh_dates`,
  `_rc_line`. Data: `HOLIDAY_NAME_MAP` (שמיני עצרת→שמחת תורה), `EVENT_HOLIDAYS`, `EVENT_FASTS` (unused/deferred).

### Rules (all user-confirmed)
- Format `שם - יום ד.מ`. Holidays = single day EXCEPT **ראש השנה** = 2-day range (candles both days).
- **שביעי של פסח** injected separately (21 Nissan) — pyluach calls all 7 Pesach days "פסח".
- Rosh Chodesh: show BOTH current-month RC and next-month RC (calendar published before 1st Shabbat →
  needs next month's start). Skip current RC if on Shabbat OR = Rosh Hashana.
- Tishrei exception: on ELUL page, next-RC (1 Tishrei) shown as "ראש השנה"; omit if it's Shabbat.
- **Fasts DEFERRED** (EVENT_FASTS ready). **Manual override DEFERRED** (automatic only).

### Verified (matches user's Tishrei example exactly)
- Tishrei: ראש השנה שלישי-רביעי 23-24.9 | יום כיפור חמישי 2.10 | סוכות שלישי 7.10 |
  שמחת תורה שלישי 14.10 | ר"ח חשון רביעי-חמישי 22-23.10. Accent color 5B2D8C ✓
- Nissan: פסח 2.4 + שביעי של פסח 8.4 (separate) ✓
- Latest drafts: draft16_tishrei / _nissan / _elul.

### KEY INSIGHT from user (explains the whole "next month" logic)
The calendar is published BEFORE the first Shabbat of the month. If RC is early in the week, much of
that week already uses the calendar → it must carry next-month basics (weekday tefila times, next RC,
next parsha). This is why: week-label lookahead, next-month RC, and the ephem mincha fallback all exist.

### NEXT (deferred)
- **Special Shabbatot** (שקלים, זכור, פרה, החודש, הגדול, שובה, נחמו...) → small-font tag NEXT TO the
  parsha in the Shabbat table (NOT the events section). User's next feature.
- Fasts in events section (enable EVENT_FASTS).
- Events manual override, if needed.
- Page-fit: Tishrei has 5 event lines (most) — verify one-page still holds with events added.

---

## 2026-08-03 SESSION (part 5) — cosmetics: theme, fonts, layout, one-page fit

### Theme system (--theme)
- `THEMES` dict + `BG_HDR()`/`ACCENT()` functions; `--theme {olive,purple,gray}` (default olive).
- **purple** (current choice): fill `D9C7E8`, accent `5B2D8C` (darkened per feedback).
- **olive** kept available: fill `E6E9D8`, accent `6B7040`. User may switch later.
- Header cell fills use `BG_HDR()`; event lines use `ACCENT()` color.

### Fonts (Hebrew szCs bug already fixed in part 3)
- Title 28, זמני השבת 22, תפילות יום חול 20.
- All table DATA = 16. All colored HEADER cells = 16 (was 10/11).
- Footer/org lines = 9 (baked into ORG_LINES; was 9/8/8/8/7).
- Bottom comment lines (order): (1) `* מניין ראשון...` size 12, (2) `דף יומי: 19:45 | ערבית 20:30 (בכל ימות השבוע)` size 18 bold.

### Structure (matched to reference PDF YAKIR-5786-11-v1)
- Logo enlarged 3.3→4.65×2.86cm; header row 2.9cm.
- Shabbat table: single tall header row; שחרית שבת = ONE merged header cell over the 8:30/6:45 data cols.
- Weekday שחרית: rebuilt 4-col constants table — grouped header (שחרית|ראש חודש), row
  `06:15 | 8:30/6:45 | 06:05 | 8:30/06:30` (constants SHACH_WK_*).
- מנחה table: **Option A** — 2 cols (שבוע|מנחה); daf/arvit2 moved to the bottom constants line.
  ⚠️ USER MAY REVERT to per-row constants (old 4-col: daf 19:45 + arvit2 20:30 each row).
  WEEKDAY_DAF/WEEKDAY_ARVIT2 constants kept for easy revert. **Awaiting user feedback on this.**
- Candles NOT bold; shacharit1 bold only when ≠ 6:45.

### One-page fit (verified by user: both 4- and 5-shabbat months fit)
- Vertical margins MAR_V=0.6cm (left/right MAR=1.0cm).
- Section gap before תפילות יום חול: 10pt. Footnote before=3, constants before=1.
- Data rows: exact=False (grow-to-fit, never clip). Shabbat rows min 0.66cm; mincha/shacharit 0.72cm.
- Worst case = 5-shabbat month (Iyar 578608 / Kislev 578603). Both fit.

### Test-file convention
- Increment draftN each round (avoids Word/Excel file-lock on overwrite). Latest: draft14.
- Test months: Elul 578612 (4-shabbat, has year-boundary row), Iyar 578608 (5-shabbat, tightest fit).

### NEXT (deferred)
- **Events section** (Rosh Chodesh + holidays as accent-colored lines like reference's `ט' אב – ...`).
  Open questions: which events, computed (pyluach) vs manual, exact format. YomText renders empty until built.
- More cosmetics pending user feedback on the constant-Arvit-times layout (Option A vs revert).
- A Word (.docx) reference from the user would let us match borders/spacing exactly.

---

## 2026-08-03 SESSION (part 4) — year-boundary lookahead (last week fixes)

### Problem
Last CSV week (05/09/2026, נצבים-וילך — last Shabbat before RH 5787) had: blank weekday mincha +
week label with no "next parsha". Both because there's no "next week" row (it's next year).

### Fix (yakir_calendar.py — generator only, no CSV change)
- **`next_shabbat_label(date)`** — next Shabbat's name via **pyluach `israel=True`** (parsha, or
  holiday name if a holiday falls on it). Used in `load_months_from_csv` when no next CSV row.
  → last label now reads `נצבים-וילך - ראש השנה`.
- **`weekday_mincha_for_week(sunday)`** — computes mincha via **ephem** (same rule as modiin_weeks.py:
  floor5(earliest Sun–Thu sunset − 12)) when the CSV has no row for the "week after". Reuses imported
  helpers from modiin_weeks.py. → last row mincha now 18:40 (was blank). Self-heals any gap, not just year-end.

### CRITICAL learning — pyluach schedule
- pyluach default = **diaspora** schedule → WRONG for Modiin (diverges up to a week; 6 mismatches in 2026).
- **Always use `israel=True`** → 41/43 match; 2 remaining are spelling variants (שלח/שלח לך, פינחס/פנחס), OK.

### Verified
- Last label `נצבים-וילך - ראש השנה` ✓ | last mincha 18:40 (= ephem cross-check) ✓
- Regression Tishrei/Cheshvan: earlier weeks still from CSV, 8 rows populated, unchanged ✓
- Latest draft: `draft8_elul_5786.docx`

---

## 2026-08-03 SESSION (part 3) — cosmetics + double-parsha fix

### Cosmetics (yakir_calendar.py)
- **Hebrew font-size bug FIXED (root cause)**: `_run()` now sets `w:szCs` (complex-script size)
  and `w:bCs`, not just `w:sz`. python-docx's `font.size` only sets `w:sz` (Latin), so ALL Hebrew
  was silently falling back to 11pt. This made every prior size setting ineffective on Hebrew.
- Title (month line) → 28 · `זמני השבת` → 22 · all body/data → 14 (was 11–13).
- Shabbat table header → single tall row (both label lines in one cell), was two rows.
- Candles (הדלקת נרות) → NO longer bold.
- Shacharit-1 → bold ONLY when ≠ 6:45 (flags the exception; e.g. Rosh Chodesh 6:30). Verified.
- Space before `תפילות יום חול` increased (4pt → 18pt).

### Double-parsha fix (parse_modiin_pdf.py) — matters for NEXT YEAR too
- **Root cause**: PDF stores double-parshiot in reversed RTL order w/ messy whitespace
  (`וילך   -נצבים` instead of `נצבים-וילך`). Affects all 6 double-parshiot/year.
- **Fix**: new `normalize_parsha()` — split on dash, strip, reverse to Torah order, clean hyphen.
  Runs automatically at parse time. Verified 6/6 correct. See RULES.md Tool 1 section.
- **Applied**: regenerated `modiin_calendar.csv` (from PDF) → `modiin_weeks.csv`. Both correct now.

### Draft
- Latest: `draft7_elul_5786.docx`. Test-file convention: use `draftN_…` (increment N) to avoid
  the "file locked in Word/Excel" overwrite error.

### OPEN / next
- Week label for a double-parsha week reads `כי תבא - נצבים-וילך` (dash between weeks + dash inside
  the pair). User OK for now; could switch the week separator to en-dash/spacing if desired.
- Still pending: YomText (special events) + Rosh Chodesh (deferred). Cosmetics review ongoing.

---

## 2026-08-03 SESSION (part 2) — generator builds months from CSV (XLS no longer required)

### What changed
- **`yakir_calendar.py` now builds month structure from `modiin_weeks.csv`** via new
  `load_months_from_csv()`. Groups weeks by `hyrmo`. The XLS is no longer needed to render a month.
- **`--source {auto,csv,xls}`** (default `auto` = CSV if present, else XLS). `--source xls` forces old path.
- **Week label (שבוע)** derived as `clean(parsha_this) - clean(parsha_next)` across the whole file
  (crosses month boundaries correctly). `clean()` strips `- מברכים`.
- **daf=19:45, arvit2=20:30** now code constants (`WEEKDAY_DAF`, `WEEKDAY_ARVIT2`) — confirmed constant
  all year. shacharit2=8:30, shacharit1=6:45 already constants.
- **YomText + Rosh Chodesh render EMPTY** (no CSV source yet — deferred).

### Verified
- **Elul 5786 (`--hyrmo 578612`) now generates** ✓ (XLS could not — it only had Tishrei/Cheshvan)
- Regression Tishrei/Cheshvan from CSV ✓ | full 13-month year renders ✓ | `--source xls` fallback ✓
- Month-boundary week labels cross correctly (Tishrei last week → בראשית - נח) ✓
- ⚠️ The very last CSV row (05/09/2026, last Shabbat before RH 5787) has blank weekday mincha —
  no next-week data exists. Expected; only affects the final Shabbat of the dataset.
- ⚠️ Parsha `וילך   -נצבים` shows raw odd spacing (double-parsha from PDF). Cosmetic — may want to clean.

### RESUME HERE
- [ ] Review the Elul 5786 draft: `/tmp/elul5786.docx` (regenerate with `python3 yakir_calendar.py --hyrmo 578612 --out elul5786.docx`)
- [ ] Decide on YomText (special events) — future manual free-text column.
- [ ] Rosh Chodesh decision (deferred; time = fixed constant if kept).
- [ ] Optional cosmetic: normalize double-parsha spacing (e.g. `וילך   -נצבים` → `וילך-נצבים`).
- [ ] Remaining: Arvit2/Daf are constants now; if they ever vary, promote to manual columns.

### Approved plan file
`/home/i548802/.claude/plans/rosy-wibbling-rainbow.md`

---

## 2026-08-02 SESSION (part 2) — full Shabbat block + Hebrew month/year + hyrmo selection

### What changed
1. **Hebrew month/year columns** in `modiin_weeks.csv` (via `pyluach`, Tishrei=1): `hyrmo`,
   `heb_year`, `heb_year_str`, `heb_month_num`, `heb_month_name`. Fixes the confusing 0-based
   `--months 12 13`. Generator now supports `--hyrmo 578601 578602` (= Tishrei & Cheshvan 5786).
2. **Full Shabbat block** ported to CSV (calc/adj pattern):
   - `taarich` (pyluach Hebrew date, e.g. ה׳ תשרי)
   - `candles`, `motzash` (from PDF)
   - `sharbat_mincha` yes/no (renamed from `shabbat_mincha` — Yair Sharbat's summer 17:15 mincha)
   - `shabbat_mincha_ktana_calc/adj` = floor5(candles)
   - mincha erev shabbat = ktana + 10 — DERIVED at render (not a column; see 08-03 note above)
   - `shacharit1` (blank → default 6:45; fill to override e.g. RC 6:30)
   - `sof_zman_shma_calc/adj` = GRA via ephem (calibrated MAE 0.36 min, horizon -0:50)
   - Constants in code (not columns): mincha gedola 13:30, Sharbat 17:15, shacharit2 8:30
3. **`yakir_calendar.py`** — new `load_shabbat_block()` + `format_shabbat_mincha()`; Shabbat table
   now reads all these from CSV (keyed by weekend_date), falls back to XLS per-field if CSV missing.
   `--hyrmo` selection added; `load_months` captures `a-hyrmo`.
4. **`load_existing_manual()`** now preserves ALL manual columns across regeneration (was adj-only).
5. Docs: `modiin_weeks.py` docstring, RULES.md §2/§5 + column table.

### Verified
- Shabbat mincha rules match reliable 5785 XLS: 50/50 weeks (ktana=floor5(candles), erev=+10) ✓
- Sof zman shma vs XLS OU: MAE 0.36 min, all ±1 ✓ | taarich vs XLS: 0 mismatch ✓
- month/year vs XLS: 0 mismatch (9 overlap) ✓ | leap-year Adar I/II correct ✓
- Rendered docx: summer/winter mincha string switches on sharbat flag ✓
- adj override + preserve-on-regen + missing-CSV fallback all ✓
- ⚠️ XLS 5786 rows (idx 12-14) are CORRUPT (candles hardcoded 19:00) — expected; that's why we migrate.
- ⚠️ Tested against `/tmp/modiin_weeks_test.csv` (real `../modiin_weeks.csv` was file-locked). **Re-run
  `python3 modiin_weeks.py` with the real file unlocked.** New dep: `pyluach`.

### OPEN QUESTION for next session
- ~~Erev-shabbat vs ktana-adj coupling~~ **RESOLVED 2026-08-03**: erev shabbat is now DERIVED at
  render time as `resolved_ktana + 10` (not a stored column). Matches the old calendar exactly
  (erev = ktana + 10, verified 50/50 weeks). A `shabbat_mincha_ktana_adj` override auto-shifts erev.
  Removed the `mincha_erev_shabbat_calc/adj` columns from the CSV. Verified: normal render erev=ktana+10;
  override 16:35→erev 16:45. ✓

### RESUME HERE
- [ ] Close `modiin_weeks.csv` in Excel, re-run `python3 modiin_weeks.py`.
- [ ] User review of the full Shabbat block across all 15 months.
- [ ] Rosh Chodesh decision (deferred — you're checking the printed calendar; time = fixed constant).
- [ ] Remaining weekday-block fields: Arvit2, Daf Yomi (manual columns), week-label (derive from names).
- [ ] YomText special events (free-text manual column).

### Approved plan file
`/home/i548802/.claude/plans/rosy-wibbling-rainbow.md`

---

## 2026-08-02 SESSION (part 1) — weekday mincha rule (calc/adj) + generator reads CSV

### Decisions this session
- **Why we're migrating off the XLS**: its daily sunset data ends 2030-12-31. Building
  self-contained ephem-based CSV sources now. XLS → validation-only.
- **Weekday mincha rule**: `floor_to_5( earliest Sun–Thu sunset − 12 min )`.
  Uses EARLIEST sunset (enough time before sunset every weekday). Offset 12 min
  (old XLS "Andy Mincha" used 15). Effect for 5786: +5 min on ~20/53 weeks, rest unchanged.
- **calc + adj pattern**: `_calc` recomputed every run; `_adj` blank manual override that wins.
  This is the template for all future ported columns.
- **Scope**: mincha columns ONLY this session. Future weekday Arvit2 / Daf Yomi → manual/constant columns.

### What changed THIS session
1. **`modiin_weeks.py`** — dropped `latest_sunset_weekday` → `earliest_sunset_weekday`
   (latest only ever fed the XLS Arvit calc, which is becoming a manual column). Renamed reserved
   `weekday_mincha` → `weekday_mincha_calc` (populated) + added `weekday_mincha_adj` (blank).
   Added `floor5()`, `WEEKDAY_MINCHA_OFFSET_MIN=12`, `load_existing_adj()`, and `--no-preserve-adj`
   (regeneration preserves manual adj by default). Full column docs in module docstring.
   **Also added Hebrew month/year columns** (`hyrmo`, `heb_year`, `heb_year_str`, `heb_month_num`,
   `heb_month_name`) computed from each week's Shabbat via **pyluach**, numbered Tishrei=1 (matches XLS).
   New dep: `pyluach`.
2. **`yakir_calendar.py`** — now reads weekday mincha from `modiin_weeks.csv` via new
   `load_weekday_mincha()` + `--weeks-csv` arg. Join key = Shabbat date **+ 1 day** (that week's
   Sunday = CSV `from`). Falls back to XLS `wkMin` if CSV missing. All other fields still from XLS.
   **Added `--hyrmo` month selection** (e.g. `--hyrmo 578601 578602` = Tishrei & Cheshvan 5786),
   replacing the confusing 0-based `--months 12 13`. `--months` still works for back-compat.
   `load_months` now captures `a-hyrmo` per month.
3. **`RULES.md`** — §2 rewritten: new rule, ephem source, join alignment, full column table
   (now incl. Hebrew month/year), hyrmo selection note, migration note.

### Verified
- CSV: header has earliest_sunset_weekday, weekday_mincha_calc, weekday_mincha_adj; adj blank; delta ±2 ✓
- Rendered docx (תשרי+חשון 5786) mincha matches CSV & is +5 vs old −15 rule ✓
- adj override renders and wins over calc ✓ | preserve-adj survives regeneration ✓ | missing-CSV fallback ✓
- ⚠️ NOTE: `../modiin_weeks.csv` was file-locked (Excel/OneDrive) during this session — tested against
  `/tmp/modiin_weeks_test.csv`. **Must re-run `python3 modiin_weeks.py` with the real file unlocked.**

### RESUME HERE
- [ ] Close `modiin_weeks.csv` in Excel, re-run `python3 modiin_weeks.py` to regenerate the real file.
- [ ] User review of weekday mincha values across all 15 months.
- [ ] **sat_sof_kriat_shma** — Sof Zman Shma (GRA). Plan: calibrate ephem SUNRISE + GRA method.
- [ ] Decide remaining missing fields (Arvit2, Daf Yomi, week-label, Shabbat mincha, friMin, shacharit,
  taarich, YomText, Rosh Chodesh) — each a CSV column or constant rule. RC confirmed still rendered.

### Approved plan file (full detail)
`/home/i548802/.claude/plans/rosy-wibbling-rainbow.md`

---

## 2026-07-31 SESSION — ephem sunset + modiin_weeks.csv

### Decision this session
- The master XLS (`yakircal master v61 ...xlsx`) is now **verification-only during development**,
  NOT a data source. Biggest gap it left: daily sunset (Modiin PDF only has Shabbat/Yom-Tov times).
- Chose **ephem** to compute sunset (astral rejected — 2–4 min off).

### What was created THIS session
1. **`calibrate_ephem.py`** — ONE-TIME dev/verification tool (NOT a monthly step).
   Compares ephem sunset vs all 59 official Modiin shkia values. Result: best setting
   **horizon `-1:00`, `pressure=0`**, lat 31.8981 / lon 35.0104 → mean bias −0.02 min,
   MAE 0.70 min, worst 1.6 min, 0 dates >2 min off. Re-run only to re-verify a new year's PDF.
2. **`modiin_weeks.py`** ✅ NEW monthly step — reads `modiin_calendar.csv`, computes Sun–Sat weeks
   + ephem sunsets, writes `../modiin_weeks.csv`. Run: `python3 modiin_weeks.py`
   (flags: `--csv`, `--out`, `--horizon -1:00`). Did NOT touch parse_modiin_pdf.py or yakir_calendar.py.
3. **`../modiin_weeks.csv`** ✅ generated, 51 week rows, fully verified:
   - sunset_delta (ephem−rabanut) range −2..+1 min, 0 weeks over ±2 ✓
   - every weekend_date=Saturday, from=Sunday, to=Thursday ✓
   - shabbat_mincha matches candles≥17:30 rule, 0 mismatches ✓
   - reserved columns blank ✓
4. **`.claude/settings.local.json`** — read permission for the parent experimental folder.

### modiin_weeks.csv columns
> ⚠️ SUPERSEDED 2026-08-02 — see the 08-02 session block above for the current 14 columns.
> `latest_sunset_weekday` → `earliest_sunset_weekday`; `weekday_mincha` → `weekday_mincha_calc` + `weekday_mincha_adj`.

`week, from, to, weekend_name, weekend_date, sat_sof_kriat_shma(RESERVED/blank),
shabbat_mincha(yes/no = candles≥17:30), sunset_rabanut(PDF), sunset_ephem, sunset_delta(min),
latest_sunset_weekday(max Sun–Thu), weekday_sunset_delta(max−min Sun–Thu, 0–5 min),
weekday_mincha(RESERVED/blank)`

### RESUME HERE — deferred, in priority order
- [ ] **sat_sof_kriat_shma** — Sof Zman Shma (GRA). Plan: calibrate ephem SUNRISE + GRA method
  against a reference, same approach as the sunset calibration.
- [ ] **weekday_mincha rule** — decide delta to subtract from `latest_sunset_weekday`
  (weekday_sunset_delta is only 0–5 min, so ONE mincha per week is justified).
- [ ] **Wire yakir_calendar.py** to read modiin_weeks.csv instead of the XLS (weekday data:
  wkLbl→weekend_name, wkMin→weekday_mincha, daf, wkArv2). NOTE: yakir_calendar.py does NOT
  do the wkMin shift in code — data must arrive already-shifted.

### Approved plan file (full detail)
`/home/i548802/.claude/plans/let-s-create-modiin-weeks-csv-rules-dazzling-conway.md`

---

## Current state of each file

### parse_modiin_pdf.py ✅ COMPLETE
Parses Modiin city candles PDF → modiin_calendar.csv
Run: `python3 parse_modiin_pdf.py --pdf "../Yakir-calendar-2025-6/Modiin-2025-6-Candels.pdf"`

### populate_xls.py ✅ COMPLETE
Writes CSV data into Detail sheet of XLS. Run BEFORE opening in Excel.
Run: `python3 populate_xls.py --csv ../modiin_calendar.csv --xls "../yakircal master v61...xlsx"`
NOTE: XLS must be CLOSED in Excel when running this.

### yakir_calendar.py ✅ FUNCTIONALLY CORRECT — cosmetics still pending
Reads ResultForWordMerge → generates Word doc.
Run: `python3 yakir_calendar.py --months 12 13 --out test.docx`

Key fixes already done:
- Row heights correct (was 1000x too large, now 0.65cm)
- RTL column order correct in both weekday tables (bidiVisual)
- מנחה times correct: uses lastmo-wkMin for first row, then prev shabbat's wkMin (shift rule)
- פינחס row removed (lastmo not shown as labeled row)
- wkMin data source: ResultForWordMerge directly (NOT sunset sheet)

### RULES.md ✅ EXISTS, partially updated
Data Flow section updated with wkMin source and 2030 validity note.
Still needs: full technical review after cosmetics are done.

## PENDING — Cosmetics in yakir_calendar.py

Waiting for user feedback on test_5786_tishri_cheshvan.docx. Known issues to fix:
1. שבוע column alignment (should be center, currently may be off)
2. Any other cosmetic issues the user reports from the test file

## PENDING — Full pipeline for 5786

The XLS ResultForWordMerge only has data through חשון תשפ״ו (index 13).
For full 5786 calendar:
1. populate_xls.py already run (51 rows written to Detail sheet)
2. User needs to open XLS in Excel → Ctrl+Alt+F9 (force recalc) → Save
3. Then ResultForWordMerge will have correct 5786 values
4. Then run yakir_calendar.py for all months

## PENDING — RULES.md final update
After cosmetics done, do final RULES.md pass.

## PENDING — Test all 15 months
After XLS recalculated in Excel.

## Key technical decisions
- Candles/Motzash: ALWAYS from XLS (which got them from Modiin PDF). Never calculated.
- Weekday מנחה: from ResultForWordMerge wkMin field DIRECTLY — NO SHIFT.
  wkMin on row N = correct mincha for row N's week label.
  lastmo is NOT shown as a labeled row (only used if needed for other purposes).
  Verified against old calendar for חשון תשפ״ו: 16:35, 16:30, 16:25, 16:20 ✓
- Sunset data source: XLS sunset sheet, valid through Dec 31 2030.
  For 2031+: extend XLS sunset sheet OR use Yeshiva.org.il (±1 min of Modiin official).
- Yeshiva.org.il comparison: 15 samples across 2 months → ±1 min max, avg 0.00.
- astral library: NOT suitable, 2-4 min off from official times.
- XLS file: `yakircal master v61 Added shkia+14 times starting 5785.xlsx`
- Logo: `yakir_logo.png` in same folder as scripts.

## Quick test command
```bash
cd "/mnt/c/yfuchs/OneDrive/Documents/Yakir-Calendar-experimental/Yakir-Calendar-creator"
python3 yakir_calendar.py --months 12 13 --out test_5786.docx
```
