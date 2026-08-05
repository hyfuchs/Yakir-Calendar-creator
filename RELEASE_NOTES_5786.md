# Yakir Calendar 5786 — Release Notes: Times, Calculation & Research

**Document purpose**: explain where every time on the calendar comes from, what is computed vs.
taken from an authoritative source, and the research/calibration behind the computed values.

**Bottom line**: candle-lighting and Havdalah are **never computed** — they come from the official
Modiin municipality PDF. Everything else is either a fixed congregation constant or computed from
sun times via a calibrated astronomical model (`ephem`), validated against the official sources.

---

## 1. Why this was rebuilt

The previous calendar was produced from a master Excel workbook. That workbook's daily sun-time data
**ends on 31 December 2030**. To keep producing the calendar beyond 2030, the pipeline was rebuilt to
compute sun times directly (no end date), with the old Excel kept only for **validation**.

New pipeline:
```
Modiin municipality PDF ─▶ parse_modiin_pdf.py ─▶ modiin_calendar.csv
                                                        │
                                                        ▼
                                        modiin_weeks.py (adds ephem sun times)
                                                        │
                                                        ▼
                                                 modiin_weeks.csv
                                                        │
                                                        ▼
                                        yakir_calendar.py ─▶ Word / PDF
```

---

## 2. Source of every time

| Time | Source | Computed? |
|---|---|---|
| הדלקת נרות (candles) | **Official Modiin PDF** | No — never |
| מוצאי שבת (Havdalah) | **Official Modiin PDF** | No — never |
| שקיעה (sunset, Shabbat) | Official Modiin PDF | No |
| Weekday sunsets | `ephem` (calibrated) | Yes |
| מנחה שבת קטנה | `floor_to_5(candles)` | Yes (from PDF candles) |
| מנחה וקבלת שבת (Fri) | mincha ktana **+ 10 min** | Yes (derived) |
| מנחה גדולה | **13:30** constant | No |
| Sharbat summer mincha | **17:15** constant (when candles ≥ 17:30) | No |
| סוף זמן קריאת שמע | GRA, via `ephem` (calibrated) | Yes |
| שחרית שבת 1 | **6:45** default (exceptions below) | No |
| שחרית שבת 2 | **8:30** constant | No |
| מנחה יום חול | `floor_to_5(earliest Sun–Thu sunset − 12 min)` | Yes |
| דף יומי | **19:45** constant | No |
| ערבית 2 | **20:30** constant | No |
| Weekday שחרית (א–ה / ר״ח) | **06:15 / 06:05 / 8:30** constants | No |
| Holiday / Rosh Chodesh / fast dates | `pyluach` (Israel schedule) | Yes |

---

## 3. Astronomical calibration (the research)

Sun times are computed with the **`ephem`** library, then calibrated so results match the official
Modiin municipality times. Two independent calibrations were performed and are baked into the code
(`modiin_weeks.py`), each with its own atmospheric-refraction "horizon" angle:

### 3a. Sunset (shkia)
- **Method**: `ephem` sunset at Modiin center (lat 31.8981, lon 35.0104), horizon **−1:00**, pressure 0.
- **Validation**: compared against the official Modiin shkia values from the PDF (see §4a for exact counts).
- **Result**: essentially unbiased (mean ≈ 0), MAE ≈ **0.7 min**, within ±1 min on nearly all dates.
- Rejected alternative: the `astral` library was 2–4 min off — not accurate enough.

### 3b. Sof Zman Kriat Shma (GRA)
- **Method**: `shma = sunrise + 3 halachic hours`, where a halachic hour = (sunset − sunrise) / 12,
  i.e. `sunrise + (sunset − sunrise)/4`. `ephem`, horizon **−0:50**.
- **Validation**: compared against the master Excel's OU-sourced "Shema Latest, GRA" values (see §4b).
- **Result**: MAE ≈ **0.4 min** — tighter than the sunset fit; every sample within ±1 min.

> The two horizons (−1:00 for sunset, −0:50 for shma) are calibrated **independently** and must not be
> merged — each is the best fit for its own quantity.

**Validity**: `ephem` has no end date, so the calendar can be produced for any future year (unlike the
old 2030-limited Excel). The Excel remains only as a cross-check.

---

## 4. Validation results (5786 data)

Exact, reproducible comparisons against the authoritative sources, on the actual 5786 data set.

### 4a. Shkia — ephem vs. official Modiin PDF (51 Shabbatot)

| Offset (ephem − PDF) | Count |
|---|---|
| −2 min | 1 |
| −1 min | 14 |
| **0 (exact)** | **16** |
| +1 min | 20 |

- **50 of 51 within ±1 minute**; a single date off by 2 min.
- Mean **+0.08 min** (essentially unbiased), MAE **0.71 min**, worst case **2 min**.

### 4b. Sof Zman Kriat Shma — ephem GRA vs. official OU values (63 Shabbatot)

| Offset (ephem − OU) | Count |
|---|---|
| **0 (exact match)** | **38** |
| −1 min | 25 |

- **38 exact to the minute; the remaining 25 are exactly 1 minute earlier.** None off by more than 1 min.
- Mean **−0.40 min**, MAE **0.40 min**, worst case **1 min**.

### 4c. Weekday Mincha — new rule (−12) vs. previous rule (−15) (53 weeks)

| Result | Count |
|---|---|
| **Identical to previous** | **33** |
| **5 minutes later** | **20** |
| Other | 0 |

- Every change is exactly +5 min later; no week moves earlier. This quantifies the rule change
  (−15 → −12): about 60% of weeks unchanged, 40% five minutes later.

> Note: an earlier calibration run over the full 59-entry PDF set reported MAE ≈ 0.70 (shkia) / 0.36
> (shma). The figures above are recomputed on the current 5786 Shabbat rows and supersede those.

---

## 5. Prayer-time rules

### Weekday Mincha (changed this year)
- **New rule**: `floor_to_5( earliest Sun–Thu sunset − 12 min )`.
- Uses the **earliest** sunset of the week so the service finishes before sunset on every weekday.
- **Change from previous years**: the old rule subtracted **15 min**; the new rule subtracts **12**.
  Net effect for 5786: ~20 of 53 weeks move 5 min later; the rest are unchanged.
- Each computed value also has a manual-override column, so any week can be hand-adjusted.

### Shabbat Mincha
- **Mincha ktana** = candle time rounded **down** to the nearest 5 min.
- **Mincha erev Shabbat** (Friday) = mincha ktana **+ 10 min** (derived, so an override to ktana shifts it).
- **Mincha gedola** = 13:30 (constant); **summer 17:15** shown when candles ≥ 17:30.

### Shacharit exceptions (Shacharit 1)
- Default **6:45**; on **Shabbat Chol HaMoed Sukkot → 6:15**, **Shabbat Chol HaMoed Pesach → 6:30**.
- Rosh Chodesh weekday shacharit uses its own constants (06:05, 8:30/06:30).

---

## 6. Calendar-date computations (`pyluach`, Israel schedule)

- **Parsha schedule**: Israel reading order (differs from diaspora by up to a week in some years).
- **Holidays & fasts** in the events section: יום כיפור, סוכות, שמחת תורה (= שמיני עצרת in Israel),
  פסח, שביעי של פסח, שבועות, חנוכה, פורים, ל״ג בעומר, ט״ו בשבט, ט״ו באב; fasts: צום גדליה, י׳ בטבת,
  תענית אסתר, י״ז בתמוז, ט׳ באב.
- **Rosh Hashana** shown as a 2-day range (candles both days); other chagim show day 1 only.
- **Rosh Chodesh**: both the current and next month's RC are shown (the calendar is published before the
  first Shabbat, so it must carry the next month's start too); a 30-day month gives a 2-day RC.
- **Four special Shabbatot** (שקלים, זכור, פרה, החודש) tagged next to the parsha.
- **Shabbat Chol HaMoed**: the PDF's confusing Torah-reading name (`ויאמר משה`) is replaced with
  `שבת חוה"מ סוכות` / `שבת חוה"מ פסח`.

---

## 7. What is authoritative vs. computed — summary

- **Authoritative (never touched)**: candle-lighting, Havdalah, and Shabbat sunset — straight from the
  Modiin municipality PDF.
- **Computed & calibrated (±<1 min vs. official)**: weekday sunsets, weekday mincha, sof zman shma.
- **Fixed congregation constants**: 13:30, 17:15, 6:45, 8:30, daf 19:45, arvit2 20:30, etc.
- **Calendar logic (pyluach, Israel)**: parshiot, holidays, fasts, Rosh Chodesh, special Shabbatot.
- Every computed prayer time has a manual-override column for one-off adjustments.

*Times are for Modiin, Israel schedule. Calendar year 5786.*
