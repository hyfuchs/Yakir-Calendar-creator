#!/usr/bin/env python3
"""
Calibrate ephem sunset against official Modiin shkia values.

Reads modiin_calendar.csv (official shkia per Shabbat/Yom Tov), computes ephem
sunset for a range of (elevation, horizon) settings, and reports which settings
best reproduce the official times.

Goal: find settings where ephem matches official shkia within ~1 min, so ephem
can be trusted for the DAILY sunsets the Modiin PDF does not provide.
"""
import csv
import datetime
import sys
from zoneinfo import ZoneInfo

import ephem

CSV = "/mnt/c/yfuchs/OneDrive/Documents/Yakir-Calendar-experimental/modiin_calendar.csv"
TZ = ZoneInfo("Asia/Jerusalem")

# Modiin (Modi'in-Maccabim-Re'ut) approximate center
LAT = 31.8981
LON = 35.0104


def load_official():
    rows = []
    with open(CSV, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            shkia = r["shkia"].strip()
            if not shkia or shkia.startswith("-"):
                continue
            d = datetime.datetime.strptime(r["date_g"].strip(), "%d/%m/%Y").date()
            hh, mm = map(int, shkia.split(":"))
            rows.append((d, hh * 60 + mm))
    return rows


def ephem_sunset(date, elevation_m, horizon_deg):
    """Return local sunset as minutes-since-midnight for the given date."""
    obs = ephem.Observer()
    obs.lat = str(LAT)
    obs.lon = str(LON)
    obs.elevation = elevation_m
    obs.horizon = str(horizon_deg)
    # pressure=0 disables ephem's own refraction model so `horizon` fully controls it
    obs.pressure = 0
    # Search from local noon so we get *this* date's sunset
    noon_local = datetime.datetime(date.year, date.month, date.day, 12, 0, tzinfo=TZ)
    obs.date = ephem.Date(noon_local.astimezone(datetime.timezone.utc).replace(tzinfo=None))
    sun = ephem.Sun()
    setting_utc = obs.next_setting(sun).datetime().replace(tzinfo=datetime.timezone.utc)
    local = setting_utc.astimezone(TZ)
    return local.hour * 60 + local.minute + local.second / 60.0


def evaluate(official, elevation_m, horizon_deg):
    diffs = []
    for d, off_min in official:
        calc = ephem_sunset(d, elevation_m, horizon_deg)
        diffs.append(calc - off_min)  # +ve => ephem later than official
    n = len(diffs)
    mean = sum(diffs) / n
    mae = sum(abs(x) for x in diffs) / n
    worst = max(diffs, key=abs)
    within1 = sum(1 for x in diffs if abs(x) <= 1.0)
    return mean, mae, worst, within1, n, diffs


def main():
    official = load_official()
    print(f"Loaded {len(official)} official shkia values "
          f"({official[0][0]} .. {official[-1][0]})\n")

    # Standard geometric sunset horizon is -0:50 (sun's upper limb + refraction).
    # We scan a grid; elevation on Modiin's hills delays visible sunset.
    print(f"{'elev(m)':>8} {'horizon':>9} {'mean':>7} {'MAE':>6} {'worst':>7} {'within±1':>9}")
    print("-" * 52)
    best = None
    elev = 0  # elevation has no effect with pressure=0; horizon is the knob
    for horizon in ("-0:48", "-0:50", "-0:52", "-0:54", "-0:56",
                    "-0:58", "-1:00", "-1:02", "-1:05"):
        mean, mae, worst, within1, n, _ = evaluate(official, elev, horizon)
        print(f"{elev:>8} {horizon:>9} {mean:>+7.2f} {mae:>6.2f} "
              f"{worst:>+7.2f} {within1:>6}/{n}")
        if best is None or mae < best[0]:
            best = (mae, elev, horizon)
    print("-" * 52)
    print(f"\nBest by MAE: elevation={best[1]}m horizon={best[2]}  (MAE={best[0]:.2f} min)")

    # Detailed per-date diff for the best setting
    mean, mae, worst, within1, n, diffs = evaluate(official, best[1], best[2])
    print(f"\nDetail for best setting — mean {mean:+.2f}, MAE {mae:.2f}, "
          f"worst {worst:+.2f}, within ±1 min: {within1}/{n}")
    over2 = [(d, dv) for (d, _), dv in zip(official, diffs) if abs(dv) > 2.0]
    if over2:
        print(f"\nDates off by >2 min ({len(over2)}):")
        for d, dv in over2:
            print(f"  {d}  {dv:+.2f} min")
    else:
        print("\nNo date off by more than 2 minutes. ✓")


if __name__ == "__main__":
    main()
