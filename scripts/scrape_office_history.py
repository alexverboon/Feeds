#!/usr/bin/env python3
"""
Scrape Microsoft 365 Apps update-history page and save/update a CSV.
Output: data/office_update_history_2018-present.csv
"""

from pathlib import Path
import re, sys, requests, pandas as pd

URL     = ("https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date")
OUTFILE = Path("data/office_update_history_2018-present.csv")
OUTFILE.parent.mkdir(parents=True, exist_ok=True)

print("Downloading update-history page…", flush=True)
html = requests.get(URL, timeout=60).text

# NEW pattern: matches “May 29, 2025 … Version 2505 (Build 18827.20128)”
pat = re.compile(
    r"(?P<month>[A-Za-z]+)\s+(?P<day>\d{1,2}),\s+(?P<year>\d{4})"
    r"[^V]*Version\s+(?P<version>\d{4})\s+\(Build\s+(?P<build>[\d.]+)\)",
    re.I
)

records = []
for m in pat.finditer(html):
    date = pd.to_datetime(
        f"{m.group('year')} {m.group('month')} {m.group('day')}"
    ).strftime("%Y-%m-%d")
    records.append(
        {"Release Date": date,
         "Version":      m.group("version"),
         "Build":        m.group("build")}
    )

if not records:
    print("ERROR: zero matches – Microsoft page layout may have changed.",
          file=sys.stderr)
    sys.exit(1)           # fail the job so you notice

df = (pd.DataFrame(records)
        .drop_duplicates()
        .sort_values("Release Date", ascending=False)
        .reset_index(drop=True))

changed = (not OUTFILE.exists()) or not df.equals(pd.read_csv(OUTFILE))
df.to_csv(OUTFILE, index=False, encoding="utf-8")
print(f"Saved {len(df)} rows → {OUTFILE}  (changed={changed})")

# Return 0 even if no change; the git step will detect differences
