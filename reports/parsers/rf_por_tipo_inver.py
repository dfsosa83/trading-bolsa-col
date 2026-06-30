"""
Parser for the 'RF-PorTipoInver' sheet.

Section boundaries are detected dynamically:
  - Header rows: rows where col C (index 2) == 'Sector'
  - Total rows:  first row after the header where col C == 'Total'

This handles BVC files where the sellers block shifts by 1-2 rows
depending on the number of instruments traded on a given day.

Returns
-------
dict with keys:
  'buyers'  : pd.DataFrame  (sectors × instruments)
  'sellers' : pd.DataFrame  (sectors × instruments)
"""

import pandas as pd

# Instruments occupy columns D onward (0-indexed starting at 3)
_COL_SECTOR      = 2
_COL_INSTR_START = 3


# ── Internal helpers ──────────────────────────────────────────────────────────

def _find_sections(rows: list[tuple]) -> list[dict]:
    """
    Locate all buyer/seller blocks by scanning for rows where col C == 'Sector'.
    Returns a list of dicts: {header_idx, first_data_idx, total_idx}  (0-indexed).
    """
    sections = []
    for i, row in enumerate(rows):
        cell = row[_COL_SECTOR] if len(row) > _COL_SECTOR else None
        if isinstance(cell, str) and cell.strip() == "Sector":
            # Find the Total row that closes this block
            total_idx = None
            for j in range(i + 1, len(rows)):
                tc = rows[j][_COL_SECTOR] if len(rows[j]) > _COL_SECTOR else None
                if isinstance(tc, str) and tc.strip() == "Total":
                    total_idx = j
                    break
            if total_idx is not None:
                sections.append({
                    "header_idx":     i,
                    "first_data_idx": i + 1,
                    "total_idx":      total_idx,
                })
    return sections


def _parse_block(rows: list[tuple], section: dict) -> pd.DataFrame:
    header = rows[section["header_idx"]]
    # Collect instrument names from col D onward (stop at first None)
    instruments = []
    for c in range(_COL_INSTR_START, len(header)):
        val = header[c]
        if val is None:
            break
        instruments.append(str(val).strip())

    records = []
    for row in rows[section["first_data_idx"] : section["total_idx"]]:
        sector = row[_COL_SECTOR] if len(row) > _COL_SECTOR else None
        if not isinstance(sector, str) or not sector.strip():
            continue
        values = [
            row[c] if c < len(row) else 0
            for c in range(_COL_INSTR_START, _COL_INSTR_START + len(instruments))
        ]
        records.append([sector.strip()] + values)

    df = pd.DataFrame(records, columns=["sector"] + instruments)
    for col in instruments:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["total"] = df[instruments].sum(axis=1)
    return df


# ── Public API ────────────────────────────────────────────────────────────────

def parse(ws_rows: list[tuple]) -> dict[str, pd.DataFrame]:
    """
    Parameters
    ----------
    ws_rows : list of tuples
        All rows from the sheet as returned by openpyxl iter_rows(values_only=True).

    Returns
    -------
    dict with keys 'buyers' and 'sellers'.
    """
    sections = _find_sections(ws_rows)
    if len(sections) < 2:
        raise ValueError(
            f"Expected 2 sections (buyers + sellers) in RF-PorTipoInver, found {len(sections)}"
        )
    return {
        "buyers":  _parse_block(ws_rows, sections[0]),
        "sellers": _parse_block(ws_rows, sections[1]),
    }
