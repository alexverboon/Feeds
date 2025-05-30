#!/usr/bin/env python3
"""
Scrape the Microsoft Learn “Update history (listed by date)” page and
write/refresh data/office_update_history_2018-present.csv
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

# ── 1 ▸  strip all HTML tags and collapse whitespace
text = re.sub(r"<[^>]+>", " ", html)          # remove tags
text = re.sub(r"\s+", " ", text)              # compress whitespace

# ── 2 ▸  regex for either “2025 May 29”  or  “May 29, 2025”
pat = re.compile(
    r"(?:"
    r"(?P<Y1>\d{4})\s+(?P<M1>[A-Z][a-z]+)\s+(?P<D1>\d{1,2})"         # 2025 May 29
    r"|"
    r"(?P<M2>[A-Z][a-z]+)\s+(?P<D2>\d{1,2}),\s+(?P<Y2>\d{4})"         # May 29, 2025
    r")"
    r".{0,200}?"                                                      # up to 200 chars (newlines, etc.)
    r"Version\s+(?P<ver>\d{4})\s+\(Build\s+(?P<build>[\d.]+)\)",
    re.S,
)

records = []
for m in pat.finditer(text):
    year  = m.group("Y1") or m.group("Y2")
    month = m.group("M1") or m.group("M2")
    day   = m.group("D1") or m.group("D2")
    date  = pd.to_datetime(f"{year} {month} {day}").strftime("%Y-%m-%d")
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
