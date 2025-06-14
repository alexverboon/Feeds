#!/usr/bin/env python3
"""
Scrape Microsoft Learn ‘update-history-…by-date’ and refresh
data/office_update_history_2018-present.csv
"""

from pathlib import Path
import re, sys, requests, pandas as pd

URL  = ("https://learn.microsoft.com/en-us/officeupdates/"
        "update-history-microsoft365-apps-by-date")
DEST = Path("data/office_update_history_2018-present.csv")
DEST.parent.mkdir(parents=True, exist_ok=True)

print("Downloading update-history page …", flush=True)
html = requests.get(
    URL,
    headers={"User-Agent": "office-history-bot/1.0"},
    timeout=60,
).text

# Strip tags → single-space text
plain = re.sub(r"<[^>]+>", " ", html)
plain = re.sub(r"\s+", " ", plain)

MONTHS = ("January|February|March|April|May|June|July|August|"
          "September|October|November|December")

pat = re.compile(
    rf"(?:"
    rf"(?P<Y1>\d{{4}})\s+(?P<M1>{MONTHS})\s+(?P<D1>\d{{1,2}})"      # 2025 May 29
    rf"|"
    rf"(?P<M2>{MONTHS})\s+(?P<D2>\d{{1,2}}),\s+(?P<Y2>\d{{4}})"      # May 29, 2025
    rf")"
    rf".{{0,200}}?"                                                  # up to 200 chars
    rf"Version\s+(?P<ver>\d{{4}})\s+\(Build\s+(?P<build>[\d.]+)\)",
    re.S,
)

records = []
for m in pat.finditer(plain):
    year  = m.group("Y1") or m.group("Y2")
    month = m.group("M1") or m.group("M2")
    day   = m.group("D1") or m.group("D2")
    try:
        date = pd.to_datetime(f"{year} {month} {day}").strftime("%Y-%m-%d")
    except ValueError:
        # Skip any malformed capture rather than crashing
        print(f"⚠️  Skipped malformed date: {year}-{month}-{day}", file=sys.stderr)
        continue
    records.append(
        {"Release Date": date,
         "Version":      m.group("ver"),
         "Build":        m.group("build")}
    )

if not records:
    sys.stderr.write("ERROR: regex found zero releases – check pattern.\n")
    sys.exit(1)

df = (pd.DataFrame(records)
        .drop_duplicates()
        .sort_values("Release Date", ascending=False)
        .reset_index(drop=True))

changed = (not DEST.exists()) or not df.equals(pd.read_csv(DEST))
df.to_csv(DEST, index=False, encoding="utf-8")
print(f"Saved {len(df)} rows → {DEST}  (changed={changed})")
