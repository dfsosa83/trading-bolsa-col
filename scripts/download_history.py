"""
Download the most recent BVC daily xlsx and save it to data/history/.
Skips the download if the file already exists (idempotent).
Run by GitHub Actions daily at 8:15am Colombia time.
"""

import io
import os
import sys
import zipfile
from pathlib import Path

import requests

# ── Reuse config from download_bvc_daily.py ───────────────────────────────────
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from download_bvc_daily import fetch_report_list, HYGRAPH_URL, BEARER_TOKEN

HISTORY_DIR = Path(__file__).resolve().parent.parent / "data" / "history"


def download_xlsx_to_history(report: dict) -> str | None:
    """Download the xlsx for *report* into HISTORY_DIR.

    Returns the saved path, or None if the file already existed.
    """
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)

    url      = report["attached"]["url"]
    filename = report["attached"]["fileName"]          # e.g. BoletinDiario_2026_06_26.zip
    date_str = report["date"]                          # e.g. 2026-06-26

    # Derive xlsx name from zip name
    xlsx_name = filename.replace(".zip", ".xlsx")
    dest      = HISTORY_DIR / xlsx_name

    if dest.exists():
        print(f"[skip] {date_str}  {xlsx_name}  (already exists)")
        return None

    print(f"[down] {date_str}  {filename}")
    response = requests.get(url, timeout=60)
    response.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
        xlsx_entries = [n for n in zf.namelist() if n.lower().endswith(".xlsx")]
        if not xlsx_entries:
            print(f"  WARNING: no xlsx found in {filename}")
            return None
        data = zf.read(xlsx_entries[0])

    dest.write_bytes(data)
    print(f"  → saved {dest.name}  ({len(data):,} bytes)")
    return str(dest)


def main():
    print("Fetching report list from BVC...")
    reports = fetch_report_list()
    print(f"  {len(reports)} reports available\n")

    saved = 0
    for report in reports:
        result = download_xlsx_to_history(report)
        if result:
            saved += 1

    print(f"\nDone. {saved} new file(s) saved to {HISTORY_DIR}")


if __name__ == "__main__":
    main()
