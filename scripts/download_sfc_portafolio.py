"""
Download and parse SFC "Portafolio de Inversión Detallado" monthly files.
Extracts BGLT bond holdings from the Formato_351 sheet and saves them to
data/sfc_portafolio/bglt_holdings.csv for use in the Streamlit dashboard.

Source : https://www.superfinanciera.gov.co/publicacion/10097246
Format : Old-style .xls (Excel 97-2003), parsed with xlrd
Run    : python scripts/download_sfc_portafolio.py

When a new month is published by SFC, add its idFile to KNOWN_FILES and
re-run this script, then commit the updated CSV.

How to find a new file ID:
  1. Open https://www.superfinanciera.gov.co/publicacion/10097246
  2. Right-click the new month's download link → Copy link address
  3. The URL ends with &idFile=XXXXXXX — add that number below.
"""

import csv
import io
import pathlib
import sys

import openpyxl
import requests
import xlrd

# ── Config ────────────────────────────────────────────────────────────────────

SFC_DOWNLOAD_URL = (
    "https://www.superfinanciera.gov.co/loader.php"
    "?lServicio=Tools2&lTipo=descargas&lFuncion=descargar&idFile={file_id}"
)

# Known monthly file IDs — update when SFC publishes new months
KNOWN_FILES = {
    "2026-01": 1080752,
    "2026-02": 1081227,
    "2026-03": 1081409,
    "2026-04": 1081895,
    "2026-05": 1082057,
    # Add new months here as SFC publishes them:
    # "2026-06": XXXXXXX,
}

_ROOT      = pathlib.Path(__file__).resolve().parent.parent
OUTPUT_DIR = _ROOT / "data" / "sfc_portafolio"
OUTPUT_CSV = OUTPUT_DIR / "bglt_holdings.csv"

# ── Column indices in Formato_351 (0-indexed, confirmed from May 2026 file) ──
COL_ENTITY = 2    # Nombre de Entidad
COL_PATRIM = 6    # Nombre Patrimonio (Moderado, Conservador, etc.)
COL_NEMO   = 27   # Nemotecnico
COL_MONEDA = 34   # Código Moneda
COL_MKT    = 53   # Vr. mercado o Vr presente en $ (mark-to-market in COP)

# ── Helpers ───────────────────────────────────────────────────────────────────

def _clean_entity(raw: str) -> str:
    """Normalize AFP names to short display names."""
    name = str(raw).strip().replace('"', '').split('"')[0].strip()
    return (name
            .replace("COLFONDOS S.A. Y COLFONDOS", "Colfondos")
            .replace("COLFONDOS S.A.", "Colfondos")
            .replace("PORVENIR", "Porvenir")
            .replace("PROTECCION", "Protección")
            .replace("SKANDIA", "Skandia"))


def parse_xls_bytes(data: bytes, month_str: str) -> list[dict]:
    """
    Parse one SFC portfolio file and return only BGLT bond rows.
    Supports both old .xls (xlrd) and modern .xlsx (openpyxl) formats —
    SFC has used both across months.

    Parameters
    ----------
    data       : raw bytes of the file
    month_str  : YYYY-MM label for this cut-off (e.g. '2026-05')

    Returns
    -------
    List of dicts with keys: fecha, entidad, fondo, nemotecnico, moneda, valor_cop
    """
    # Detect format from magic bytes: PK = zip/xlsx, D0CF = old xls
    is_xlsx = data[:2] == b"PK"

    if is_xlsx:
        wb_opx = openpyxl.load_workbook(
            io.BytesIO(data), read_only=True, data_only=True
        )
        sh_opx = wb_opx["Formato_351"]
        rows_iter = sh_opx.iter_rows(values_only=True)
        next(rows_iter)  # skip header row

        rows = []
        for row in rows_iter:
            if len(row) <= max(COL_NEMO, COL_MKT):
                continue
            nemo = str(row[COL_NEMO] or "").strip()
            if "BGLT" not in nemo:
                continue
            rows.append({
                "fecha":       month_str,
                "entidad":     _clean_entity(row[COL_ENTITY] or ""),
                "fondo":       str(row[COL_PATRIM] or "").strip(),
                "nemotecnico": nemo,
                "moneda":      str(row[COL_MONEDA] or "").strip(),
                "valor_cop":   round(float(row[COL_MKT] or 0), 2),
            })
        wb_opx.close()
        return rows

    else:
        wb = xlrd.open_workbook(file_contents=data)
        sh = wb.sheet_by_name("Formato_351")

        rows = []
        for r in range(1, sh.nrows):   # row 0 is the header
            nemo = str(sh.cell_value(r, COL_NEMO)).strip()
            if "BGLT" not in nemo:
                continue
            rows.append({
                "fecha":       month_str,
                "entidad":     _clean_entity(sh.cell_value(r, COL_ENTITY)),
                "fondo":       str(sh.cell_value(r, COL_PATRIM)).strip(),
                "nemotecnico": nemo,
                "moneda":      str(sh.cell_value(r, COL_MONEDA)).strip(),
                "valor_cop":   round(float(sh.cell_value(r, COL_MKT) or 0), 2),
            })
        return rows


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    all_rows: list[dict] = []
    errors: list[str] = []

    for month, file_id in sorted(KNOWN_FILES.items()):
        url = SFC_DOWNLOAD_URL.format(file_id=file_id)
        print(f"[{month}] Downloading idFile={file_id} ...", end=" ", flush=True)
        try:
            resp = requests.get(url, timeout=90)
            resp.raise_for_status()
            rows = parse_xls_bytes(resp.content, month)
            print(f"{len(rows)} BGLT rows  ({len(resp.content)/1e6:.1f} MB)")
            all_rows.extend(rows)
        except Exception as exc:
            print(f"ERROR — {exc}")
            errors.append(f"{month}: {exc}")

    if not all_rows:
        print("\nNo data extracted. Check errors above.")
        sys.exit(1)

    # Write CSV
    fieldnames = ["fecha", "entidad", "fondo", "nemotecnico", "moneda", "valor_cop"]
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    months = sorted({r["fecha"] for r in all_rows})
    print(f"\n✅ Saved {len(all_rows):,} rows  →  {OUTPUT_CSV}")
    print(f"   Months: {months[0]} → {months[-1]}")
    print(f"   Entities: {sorted({r['entidad'] for r in all_rows})}")
    print(f"   BGLT bonds: {sorted({r['nemotecnico'] for r in all_rows})}")

    if errors:
        print(f"\n⚠️  {len(errors)} error(s):")
        for e in errors:
            print(f"   {e}")


if __name__ == "__main__":
    main()
