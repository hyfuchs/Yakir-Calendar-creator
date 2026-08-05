# How to Run — Yakir Calendar Creator

Generates the congregation's monthly prayer-time calendar (one A4-landscape page per
Hebrew month) as a Word `.docx` and optional `.pdf`.

**The pipeline is now fully CSV-driven.** The old master Excel is no longer used to
generate the calendar (kept only as an optional validation fallback). Sun times are
computed with the `ephem` library, calibrated to the official Modiin municipality times —
so the calendar works for any year, not just those in the old 2030-limited Excel.

---

## 0. One-time setup (any machine)

Requires **Python 3.9+**. Install the dependencies:

```bash
pip install -r requirements.txt
```

(Full per-OS install instructions — WSL / Windows / Mac — are in **INSTALL.md**.)

For **PDF output** you also need either LibreOffice or Microsoft Word installed
(see INSTALL.md). The `.docx` is always produced regardless.

---

## Producing a NEW YEAR (e.g. תשפ״ז) — 3 steps

You need the new year's **official Modiin municipality candles PDF** (the annual
one-page table of candle-lighting / shkia / motzash).

> **The PDF must be a DIGITAL PDF (has a selectable/searchable text layer)** — that's what
> the parser reads. A **scanned image** (photo/scan saved as PDF, text not selectable) will
> yield **0 rows**. If you only have a scan, skip Steps 1–2 and use the
> **[scanned-image workflow](#when-you-only-have-a-scanned-image-no-digital-pdf)** below instead.

### Step 1 — Parse the Modiin PDF → `modiin_calendar.csv`

```bash
cd <project folder>/Yakir-Calendar-creator
python3 parse_modiin_pdf.py --pdf "path/to/Modiin-<year>-Candels.pdf"
```

Handles automatically: double-parshiot ordering, candles/shkia/motzash extraction.
Writes `modiin_calendar.csv` **next to the input PDF** (the default when `--out` is omitted) — so
running it from a year folder puts the CSV in that folder. Pass `--out` to choose a different path.
Expected: `CSV saved → …/modiin_calendar.csv (NN rows)`.

> ⚠️ If the new PDF has a different column layout, the x-coordinate ranges in
> `parse_modiin_pdf.py` (COL_* constants) may need adjusting. Check the parsed CSV looks
> right (dates, candle times) before continuing. Spot-check a few Shabbatot against the PDF.

### Step 2 — Build the week data → `modiin_weeks.csv`

```bash
python3 modiin_weeks.py --csv "path/to/modiin_calendar.csv"
```

Adds ephem sun times, Hebrew month/year, weekday mincha, and the full Shabbat block.
Writes `modiin_weeks.csv` **next to the input `--csv`** (default when `--out` is omitted).
Expected: `Wrote NN week rows` + `✓ all weeks within ±2 min of official shkia`.
The ±2 min check is your validation that ephem still matches the official times.

> Manual overrides you typed into `modiin_weeks.csv` (the `_adj` columns) are preserved
> across re-runs. Use `--no-preserve-adj` to reset them.

### Step 3 — Generate the calendar

Pass the Hebrew year with `--year` — it generates **all** the year's months in order,
figuring out 12-vs-13 (leap year) automatically from the data:

```bash
python3 yakir_calendar.py --theme purple --pdf \
  --weeks-csv "path/to/modiin_weeks.csv" \
  --year 5787
```

With `--out` omitted, writes `yakir_calendar_output.docx` (+ `.pdf`) **next to the `--weeks-csv`**.
Pass `--out yakir_5787.docx` to name it. Open and review.

> **Tip — run from the year folder.** `cd` into `Yakir-calendar-<year>/` and run each tool. Inputs
> default to the **current folder** and outputs land there too — so you can omit the input/output paths
> entirely:
> ```bash
> cd ../Yakir-calendar-2026-7
> python3 ../Yakir-Calendar-creator/parse_modiin_pdf.py --pdf Modiin-2026-7-Candels.pdf
> python3 ../Yakir-Calendar-creator/modiin_weeks.py
> python3 ../Yakir-Calendar-creator/yakir_calendar.py --theme purple --pdf --year 5787 --out yakir_5787.docx
> ```
> (`modiin_weeks.py` reads `modiin_calendar.csv` from the current folder; `yakir_calendar.py` reads
> `modiin_weeks.csv` from it. Only `--year` and an optional `--out` name are worth passing.)

> **Leap years handled for you.** `--year` selects whatever months exist in the CSV — 12 for a normal
> year, **13 for a leap year** (e.g. תשפ״ז/5787, with אדר א׳ + אדר ב׳). You don't need to know or count.
> If the CSV doesn't fully cover the year, it prints a warning telling you how many months it found.
>
> **Selecting individual months.** To render just one or a few months (e.g. for a quick review), use
> `--hyrmo` with year+month keys (Tishrei=01), e.g. `--hyrmo 578701 578702`. Running with neither
> `--year` nor `--hyrmo` outputs *all* rows in the CSV, including a partial trailing month from the
> previous year — so always pass one of them for a clean result.

---

## When you only have a SCANNED IMAGE (no digital PDF)

Some years the municipality table is only available as an image scan (no text to parse). Everything
except the **times** is computable, so generate a skeleton and type in the times:

### Step 1 — Generate a skeleton CSV (dates + parshiot pre-filled, times blank)

```bash
python3 make_skeleton_csv.py --year 5788 --out "../Yakir-calendar-<folder>/modiin_calendar.csv"
```

Fills `date_g`, `date_h`, `parsha`, `holiday` from pyluach (Israel schedule; handles leap years and
double-parshiot). Leaves `candles`, `shkia`, `motzash` **blank**, and adds `candles_hint`/`shkia_hint`/
`motzash_hint` columns with ephem estimates.

### Step 2 — Type the official times from the scan

Open the CSV in Excel. For each Saturday, type the **candles** and **מוצאי שבת (motzash)** from the scan
(and **shkia** if you want the ±2-min cross-check). Use the `*_hint` columns to sanity-check your typing
— your official value should be within a couple minutes of the hint. The `_hint` columns are ignored by
the pipeline; delete them or leave them. Save.

### Step 3 — Continue as normal

```bash
python3 modiin_weeks.py --csv "…/modiin_calendar.csv" --out "…/modiin_weeks.csv"
python3 yakir_calendar.py --weeks-csv "…/modiin_weeks.csv" --year <year> --out "…/yakir_<year>.docx" --pdf
```

> Note: the `parsha`/`date_h`/`holiday` columns in the skeleton are for your readability only — the
> generator recomputes parshiot & holidays from pyluach at render time. Only the **times** you type are
> actually used from the CSV.

---

## What updates automatically each year (no code changes)

Computed from the new PDF + pyluach, so they just work for תשפ״ז and beyond:
candle-lighting & motzash (from PDF), sunsets/shma/mincha (ephem), parshiot & Israel
reading schedule, holidays, fasts, Rosh Chodesh, four special Shabbatot, Chol HaMoed
relabeling, the school-vacation months (21 Jun–31 Aug), and the Chanukah note (Kislev).

**Things to eyeball on a new year** (rare, but possible): a brand-new double-parsha
combination the parser hasn't seen; a PDF layout change (Step 1 warning above);
leap years (13 months — pyluach handles Adar I/II, and `--year` picks up all 13 automatically).

---

## Common options

```bash
# Whole year (12 months, or 13 in a leap year — figured out automatically):
python3 yakir_calendar.py --year 5787 --out yakir_5787.docx --pdf

# One or a few months only (for quick review):
python3 yakir_calendar.py --hyrmo 578701 578702 --out test.docx

# Color theme (default olive; purple is the congregation's current choice):
python3 yakir_calendar.py --theme olive  ...
python3 yakir_calendar.py --theme purple ...

# Skip PDF (docx only):  omit --pdf
# Force the legacy XLS source instead of the CSV:  --source xls  (rarely needed)
```

---

## Files at a glance

| What | Where |
|---|---|
| City candles PDF (input) | the annual Modiin municipality PDF |
| Parsed calendar CSV | `modiin_calendar.csv` (project folder) |
| Week data CSV (main source) | `modiin_weeks.csv` (project folder) |
| Scripts | `Yakir-Calendar-creator/` (parse_modiin_pdf.py, modiin_weeks.py, yakir_calendar.py) |
| Logo | `Yakir-Calendar-creator/yakir_logo.png` |
| Output | `Yakir-Calendar-creator/yakir_<year>.docx` / `.pdf` |
| Rules & technical reference | `RULES.md` |
| Times/calculation research | `RELEASE_NOTES_5786.md` |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `PermissionError` writing a file | The .docx/.csv is open in Word/Excel — close it, or output to a new filename. |
| PDF not produced | LibreOffice/Word not found — open the .docx and Save As → PDF manually. See INSTALL.md. |
| Hebrew renders as boxes | Missing Hebrew fonts — install system fonts (INSTALL.md). |
| A month spills to 2 pages | A long week label wrapped; the שבוע column is sized for the worst case, but check RULES.md layout notes. |
| Wrong parsha for a Shabbat | Double-parsha / Chol HaMoed handling — see RULES.md §Tool 1 and §Shabbat Chol HaMoed. |
