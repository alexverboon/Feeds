#!/usr/bin/env python3
"""
Scrape Microsoft 365 Apps update-history-by-date page and save/refresh a CSV.

Creates/overwrites: data/office_update_history_2018-present.csv
"""

from pathlib import Path
import re, requests, pandas as pd
from bs4 import BeautifulSoup

URL = (
    "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date"
)
OUTFILE = Path("data/office_update_history_2018-present.csv")
OUTFILE.parent.mkdir(parents=True, exist_ok=True)

print("Downloading update-history page…")
html = requests.get(URL, timeout=60).text
soup = BeautifulSoup(html, "html.parser")

# Regex matches: "May 29, 2025 Version 2505 (Build 18827.20128)"
pattern = re.compile(
    r"(?P<month>[A-Z][a-z]+)\s+(?P<day>\d{1,2}),\s+(?P<year>\d{4})"
    r".*?Version\s+(?P<version>\d{4})\s+\(Build\s+(?P<build>[\d\.]+)\)",
    re.S,
)

records = []
for text in soup.stripped_strings:  # iterate text nodes
    m = pattern.search(text)
    if m:
        date = pd.to_datetime(
            f"{m.group('year')}-{m.group('month')}-{m.group('day')}"
        ).strftime("%Y-%m-%d")
        records.append(
            {"Release Date": date,
             "Version": m.group("version"),
             "Build": m.group("build")}
        )

df = (
    pd.DataFrame(records)
    .drop_duplicates()
    .sort_values("Release Date", ascending=False)
    .reset_index(drop=True)
)

if OUTFILE.exists():
    old = pd.read_csv(OUTFILE)
    changed = not df.equals(old)
else:
    changed = True

df.to_csv(OUTFILE, index=False, encoding="utf-8")

print(f"Saved {len(df)} rows → {OUTFILE}")
print("Changed:", changed)
exit(0 if changed else 78)  # 78 = no-change (used only for info)
