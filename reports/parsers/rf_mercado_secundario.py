"""
Parser for the 'RF-Mercado Secundario' sheet.

Sheet layout (0-indexed rows internally, but shown as 1-indexed in comments):
  Row  5  : Section header "Segmento Privado - Sistema Transaccional"
  Row  9  : Column header for Privado section
  Row 10  : Total row for Privado (no detail rows available)

  Row 12  : Section header "Segmento Público - Sistema Transaccional"
  Row 14  : Sub-header group labels  (Compra-Ventas | Simultáneas)
  Row 15  : Sub-header rate/price labels
  Row 16  : Full column header row
  Rows 17-106 : Bond data rows
  Row 107 : Total row

  Row 109 : Section header "Segmento Público - Sistema de Registro"
  Row 113 : Full column header row
  Rows 114-265 : Bond data rows
  Row 266 : Total row

Returns
-------
dict with keys:
  'pub_transaccional'  : pd.DataFrame
  'pub_registro'       : pd.DataFrame
"""

import pandas as pd


# ── Column positions (0-indexed) ──────────────────────────────────────────────
_COL_NEMO       = 2
_COL_DESC       = 3
_COL_FEC_EMIS   = 4
_COL_FEC_VCTO   = 5

# Compra-Ventas
_COL_CV_MONTO   = 6
_COL_CV_NOPES   = 7
_COL_CV_TMIN    = 8
_COL_CV_TMAX    = 9
_COL_CV_TMED    = 10
_COL_CV_TCLS    = 11
_COL_CV_PMIN    = 12
_COL_CV_PMAX    = 13
_COL_CV_PMED    = 14
_COL_CV_PCLS    = 15

# Simultáneas
_COL_SIM_MONTO  = 16
_COL_SIM_NOPES  = 17
_COL_SIM_TMIN   = 18
_COL_SIM_TMAX   = 19
_COL_SIM_TMED   = 20
_COL_SIM_TCLS   = 21
_COL_SIM_PMIN   = 22
_COL_SIM_PMAX   = 23
_COL_SIM_PMED   = 24
_COL_SIM_PCLS   = 25

_COLUMNS = [
    "nemotecnico", "descripcion", "fec_emision", "fec_vcto",
    "cv_monto", "cv_num_opes",
    "cv_tasa_min", "cv_tasa_max", "cv_tasa_med", "cv_tasa_cierre",
    "cv_precio_min", "cv_precio_max", "cv_precio_med", "cv_precio_cierre",
    "sim_monto", "sim_num_opes",
    "sim_tasa_min", "sim_tasa_max", "sim_tasa_med", "sim_tasa_cierre",
    "sim_precio_min", "sim_precio_max", "sim_precio_med", "sim_precio_cierre",
]

_SLOT = [
    _COL_NEMO, _COL_DESC, _COL_FEC_EMIS, _COL_FEC_VCTO,
    _COL_CV_MONTO, _COL_CV_NOPES,
    _COL_CV_TMIN, _COL_CV_TMAX, _COL_CV_TMED, _COL_CV_TCLS,
    _COL_CV_PMIN, _COL_CV_PMAX, _COL_CV_PMED, _COL_CV_PCLS,
    _COL_SIM_MONTO, _COL_SIM_NOPES,
    _COL_SIM_TMIN, _COL_SIM_TMAX, _COL_SIM_TMED, _COL_SIM_TCLS,
    _COL_SIM_PMIN, _COL_SIM_PMAX, _COL_SIM_PMED, _COL_SIM_PCLS,
]

# ── Section boundaries (1-indexed row numbers) ─────────────────────────────────
_SEC_PUB_TRANS_FIRST_DATA = 17    # first bond data row
_SEC_PUB_TRANS_TOTAL      = 107   # "Total" row — excluded from data

_SEC_PUB_REG_FIRST_DATA   = 114
_SEC_PUB_REG_TOTAL        = 266


# ── Internal helpers ──────────────────────────────────────────────────────────

def _is_data_row(row: tuple) -> bool:
    """True when column C (nemotécnico) is a non-empty string that isn't a label."""
    v = row[_COL_NEMO]
    return isinstance(v, str) and v.strip() not in ("", "Total", "Nemotécnico")


def _to_df(rows: list[tuple], first: int, last_excl: int) -> pd.DataFrame:
    """Extract bond rows from `first` to `last_excl` (1-indexed, exclusive end)."""
    records = []
    for row in rows[first - 1 : last_excl - 1]:
        if not _is_data_row(row):
            continue
        records.append([row[c] if c < len(row) else None for c in _SLOT])
    df = pd.DataFrame(records, columns=_COLUMNS)
    df["descripcion"] = df["descripcion"].str.strip()
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
    dict with keys 'pub_transaccional' and 'pub_registro'.
    """
    return {
        "pub_transaccional": _to_df(
            ws_rows,
            _SEC_PUB_TRANS_FIRST_DATA,
            _SEC_PUB_TRANS_TOTAL,
        ),
        "pub_registro": _to_df(
            ws_rows,
            _SEC_PUB_REG_FIRST_DATA,
            _SEC_PUB_REG_TOTAL,
        ),
    }
