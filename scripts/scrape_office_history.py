#!/usr/bin/env python3
"""
Fetch Microsoft 365 Apps (Office) build history from the official
OfficeDocs-OfficeUpdates repo and write/update CSV:
data/office_update_history_2018-present.csv
"""

from pathlib import Path
import re, sys, requests, pandas as pd

RAW_MD = (
    "https://raw.githubusercontent.com/"
    "MicrosoftDocs/OfficeDocs-OfficeUpdates/main/OfficeUpdates/"
    "update-history-microsoft365-apps-by-date.md"
)
OUT = Path("data/office_update_history_2018-present.csv")
OUT.parent.mkdir(parents=True, exist_ok=True)

print("Downloading raw markdown …", flush=True)
resp = requests.get(
    RAW_MD,
    headers={"User-Agent": "office-history-bot/1.0"},
    timeout=60,
)
resp.raise_for_status()
md = resp.text

# Regex for a markdown bullet like:
# * **May 29, 2025** – Version 2505 (Build 18827.20128)
pat = re.compile(
    r"\*\s+\*\*(?P<month>[A-Za-z]+)\s+(?P<day>\d{1,2}),\s+(?P<year>\d{4})\*\*"
    r"\s+–\s+Version\s+(?P<version>\d{4})\s+\(Build\s+(?P<build>[\d.]+)\)"
)

records = []
for m in pat.finditer(md):
    date = pd.to_datetime(
        f"{m['year']} {m['month']} {m['day']}"
    ).strftime("%Y-%m-%d")
    records.append(
        {"Release Date": date,
         "Version":      m['version'],
         "Build":        m['build']}
    )

if not records:
    print("ERROR: regex found zero releases – check pattern.", file=sys.stderr)
    sys.exit(1)

df = (
    pd.DataFrame(records)
      .drop_duplicates()
      .sort_values("Release Date", ascending=False)
      .reset_index(drop=True)
)

changed = (not OUT.exists()) or not df.equals(pd.read_csv(OUT))
df.to_csv(OUT, index=False, encoding="utf-8")
print(f"Saved {len(df)} rows → {OUT}  (changed={changed})")
