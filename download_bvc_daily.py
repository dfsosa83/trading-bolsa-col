"""
BVC Daily Reports Downloader
Fetches and downloads "Informe Diario" files from the BVC (Bolsa de Valores de Colombia)
website via their Hygraph GraphQL API. After each download the .xlsx inside the ZIP is
extracted to data/daily_data/xlsx/ for further analysis.

Usage:
    python download_bvc_daily.py
"""

import io
import os
import zipfile

import requests

# ── Config ────────────────────────────────────────────────────────────────────

HYGRAPH_URL = (
    "https://us-east-1-bolsa-co.cdn.hygraph.com"
    "/content/ckdolgg6k07rc01xnc22d25r1/master"
)

# Static Bearer token embedded in BVC's public JS bundle (issued Dec 2021)
BEARER_TOKEN = (
    "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6ImdjbXMtbWFpbi1wcm9kdWN0aW9u"
    "In0.eyJ2ZXJzaW9uIjozLCJpYXQiOjE2Mzk2ODQ5NDUsImF1ZCI6WyJodHRwczovL2FwaS11"
    "cy1lYXN0LTEtYm9sc2EtY28uZ3JhcGhjbXMuY29tL3YyL2NrZG9sZ2c2azA3cmMwMXhuYzIy"
    "ZDI1cjEvbWFzdGVyIiwiaHR0cHM6Ly9tYW5hZ2VtZW50LW5leHQuZ3JhcGhjbXMuY29tIl0s"
    "ImlzcyI6Imh0dHBzOi8vbWFuYWdlbWVudC5ncmFwaGNtcy5jb20vIiwic3ViIjoiMmNmMzY5"
    "ZTMtNWYwMS00YjUwLThhNGMtZTM3YmQxNDI4ZGYwIiwianRpIjoiY2tpcDdjaWhxODdsNDAx"
    "ejEwcHplOTI4byJ9.ZFDpgyyGMnav1J6BRVt_ZT43vLFzQgf-iWddP2BmzgjCm1-zx_qZFMIR"
    "Mo0q7TRqDgFWhpQIt2Xuku0ens8KfvdkaPpnkeqzeZYufXeIjpXgyF_nF0tXsZjN3eSm-LkLY"
    "QK65dcJsA1UwJJjUu9wk1i-sDjnBU2LiXhFc6XCMXoH912cG3eb9sCwWatodMrkrV4qgK-zdK"
    "U2nc3FGJwL2X-lUtwdvnmKhRmLbnAzTmi4pLlog7KWYof5Syk44ysYF3stThRb8uJA580Wgcw"
    "7WEFOSJerpvdvnGPxHNYudSJdxQ2kF4SchJUnohUNdrSkmAgkogtWy-Vs1uB7Px03b4Qhjjx"
    "YMTnocoTY_0a7r1pulvi0vtp_foOD-XgPlR1qoza6g9LW1KLJ39gBVv6SD_TLZ4j94HvR3ilo"
    "XN8iIta_KGS4zXHg10ay6ZI1SXvomSbHxaOzZXclk1yWZDv76n1ZbwDWDfRkgf4figWLKqJC"
    "enL-_uIdXBe3lnD8odBLDqvQ_BM-We2MrCAPWzGC2Vo_fIpyYYDsS8h4AUZw3m3E7w1j_61f"
    "FJOIIRDpmSA49KA9Dy4hlO-eLQVvTPUWw-Tn6QAfZDrc-slr_EgjTMJg0KCP4e3NYlDPxWmbQ"
    "KUzErZj_GwDp85ezaCJTFapCy1A2Ax1CVirdk88TbjoqJE"
)

# Category ID for "Informes Bursátiles > Diario"
DIARIO_CATEGORY_ID = "ckx6mz2n444sx0b26lbmd4mzn"

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "data", "daily_data")
XLSX_DIR = os.path.join(os.path.dirname(__file__), "data", "daily_data", "xlsx")

# ── GraphQL Query ─────────────────────────────────────────────────────────────

QUERY = """
query ReportsAndBulletinsByCategoryAndDate($categoryId: ID!, $locales: [Locale!]!) {
  pdReportsAndBulletins(
    where: { category: { id: $categoryId } }
    stage: PUBLISHED
    locales: $locales
    orderBy: date_DESC
    first: 1000
  ) {
    date
    title
    description
    attached {
      fileName
      url
    }
  }
}
"""

# ── Functions ─────────────────────────────────────────────────────────────────

def fetch_report_list() -> list[dict]:
    """Query Hygraph GraphQL API and return list of report records."""
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {BEARER_TOKEN}",
        "Accept": "application/graphql-response+json, application/json",
    }
    payload = {
        "query": QUERY,
        "variables": {
            "categoryId": DIARIO_CATEGORY_ID,
            "date": None,
            "first": None,
            "locales": ["es_CO"],
        },
        "operationName": "ReportsAndBulletinsByCategoryAndDate",
    }
    response = requests.post(HYGRAPH_URL, json=payload, headers=headers, timeout=30)
    response.raise_for_status()
    data = response.json()

    errors = data.get("errors")
    if errors:
        raise RuntimeError(f"GraphQL errors: {errors}")

    return data["data"]["pdReportsAndBulletins"]


def download_file(url: str, dest_path: str) -> None:
    """Stream-download a file from url to dest_path."""
    with requests.get(url, stream=True, timeout=60) as r:
        r.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)


def download_xlsx_bytes(url: str) -> bytes:
    """Download a ZIP from *url* and return the first .xlsx inside as raw bytes.
    No files are written to disk — suitable for in-memory / cloud use."""
    response = requests.get(url, timeout=60)
    response.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
        xlsx_entries = [n for n in zf.namelist() if n.lower().endswith(".xlsx")]
        if not xlsx_entries:
            raise ValueError("No .xlsx file found in the downloaded ZIP")
        return zf.read(xlsx_entries[0])


def extract_xlsx(zip_path: str) -> str | None:
    """Extract the first .xlsx found in zip_path into XLSX_DIR.

    Returns the destination path if extracted, None if already present.
    """
    os.makedirs(XLSX_DIR, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zf:
        xlsx_entries = [n for n in zf.namelist() if n.lower().endswith(".xlsx")]
        if not xlsx_entries:
            return None
        entry = xlsx_entries[0]
        # Use only the bare filename (strip folder prefix inside ZIP)
        bare_name = os.path.basename(entry)
        dest = os.path.join(XLSX_DIR, bare_name)
        if os.path.exists(dest):
            return None  # already extracted
        data = zf.read(entry)
        with open(dest, "wb") as f:
            f.write(data)
        return dest


def main() -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(XLSX_DIR, exist_ok=True)

    print("Fetching report list from BVC...")
    reports = fetch_report_list()
    print(f"  {len(reports)} reports found.\n")

    downloaded = 0
    skipped = 0
    extracted = 0

    for report in reports:
        date = report["date"]
        file_name = report["attached"]["fileName"]
        file_url = report["attached"]["url"]
        description = (report.get("description") or "").strip()

        dest = os.path.join(OUTPUT_DIR, file_name)

        if os.path.exists(dest):
            print(f"  [skip]  {date}  {file_name}  (already exists)")
            skipped += 1
        else:
            print(f"  [down]  {date}  {file_name}")
            if description:
                print(f"          {description}")
            download_file(file_url, dest)
            downloaded += 1

        # Always try to extract xlsx (idempotent — skips if already present)
        xlsx_dest = extract_xlsx(dest)
        if xlsx_dest:
            print(f"          → extracted  {os.path.basename(xlsx_dest)}")
            extracted += 1

    print(f"\nDone.")
    print(f"  ZIPs downloaded : {downloaded}  |  skipped: {skipped}")
    print(f"  XLSXs extracted : {extracted}")
    print(f"  ZIP  folder : {OUTPUT_DIR}")
    print(f"  XLSX folder : {XLSX_DIR}")


if __name__ == "__main__":
    main()
