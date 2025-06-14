#!/usr/bin/env python3
"""
Scrape Microsoft Learn update-history-microsoft365-apps-by-date
and refresh data/office_update_history_2018-present.csv
"""

from pathlib import Path
import sys, requests, pandas as pd
from bs4 import BeautifulSoup
import re

URL  = ("https://learn.microsoft.com/en-us/officeupdates/"
        "update-history-microsoft365-apps-by-date")
DEST = Path("data/office_update_history_2018-present.csv")
DEST.parent.mkdir(parents=True, exist_ok=True)

print("Downloading update-history page …", flush=True)
res = requests.get(URL, headers={"User-Agent": "office-history-bot/1.0"}, timeout=60)
soup = BeautifulSoup(res.text, "html.parser")

records = []

for table in soup.find_all("table"):
    headers = [th.text.strip() for th in table.find_all("th")]
    rows = table.find("tbody").find_all("tr") if table.find("tbody") else table.find_all("tr")[1:]

    for row in rows:
        cells = row.find_all(["td", "th"])
        if len(cells) < 2:
            continue

        year = cells[0].text.strip()
        date = cells[1].text.strip()

        try:
            release_date = pd.to_datetime(f"{year} {date}").strftime("%Y-%m-%d")
        except Exception:
            continue  # skip rows that don’t start with a proper year + date

        for col_idx, cell in enumerate(cells[2:], start=2):
            channel = headers[col_idx] if col_idx < len(headers) else f"Column{col_idx}"
            matches = re.findall(r"Version\s+(\d+)\s*\(Build\s+([\d.]+)\)", cell.text)
            for version, build in matches:
                records.append({
                    "Release Date": release_date,
                    "Channel": channel,
                    "Version": version,
                    "Build": build
                })

if not records:
    print("ERROR: No version/build records found", file=sys.stderr)
    sys.exit(1)

df = (pd.DataFrame(records)
        .drop_duplicates()
        .sort_values(["Release Date", "Channel"], ascending=[False, True])
        .reset_index(drop=True))

changed = (not DEST.exists()) or not df.equals(pd.read_csv(DEST))
df.to_csv(DEST, index=False, encoding="utf-8")
print(f"Saved {len(df)} rows → {DEST}  (changed={changed})")
