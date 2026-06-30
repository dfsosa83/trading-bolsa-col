"""
BVC Fixed Income Dashboard
Data is fetched live from the BVC API on every session (cached 1 h).
No local files required — works on Streamlit Community Cloud.
"""

import io
import sys
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent))
from download_bvc_daily import fetch_report_list, download_xlsx_bytes
from reports.parsers import rf_mercado_secundario, rf_por_tipo_inver
from reports.generate_report import generate_to_bytes

# ── Config ─────────────────────────────────────────────────────────────────────

_EXT_DEBT_INSTR = "PUBLIC EXTERNAL DEBT BONDS DOLLAR DENOMINATED"
_HISTORY_DIR    = Path(__file__).resolve().parent / "data" / "history"

st.set_page_config(
    page_title="BVC · Bonos Ext. USD",
    page_icon="📊",
    layout="wide",
)

# ── Cached data functions ──────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def _get_reports() -> list[dict]:
    """Fetch available daily reports from BVC API. Re-checked every hour."""
    return fetch_report_list()


def _parse(xlsx_bytes: bytes, report_date: str):
    """Parse xlsx bytes into all dashboard data. Shared by both cached loaders."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ms_rows  = list(wb["RF-Mercado Secundario"].iter_rows(values_only=True))
    pti_rows = list(wb["RF-PorTipoInver"].iter_rows(values_only=True))
    wb.close()

    ms_data  = rf_mercado_secundario.parse(ms_rows)
    pti_data = rf_por_tipo_inver.parse(pti_rows)

    # ── BGLT bonds ────────────────────────────────────────────────────────────
    df_reg = ms_data["pub_registro"].copy()
    mask   = df_reg["descripcion"].str.lower().str.contains("ext dolares", na=False)
    df_bonds = df_reg[mask].copy()
    df_bonds["total_monto"] = df_bonds["cv_monto"].fillna(0) + df_bonds["sim_monto"].fillna(0)

    bond_display = df_bonds[[
        "total_monto", "nemotecnico", "fec_vcto",
        "cv_monto", "cv_num_opes",
        "cv_tasa_min", "cv_tasa_max", "cv_tasa_cierre",
    ]].rename(columns={
        "total_monto":    "Monto Total",
        "nemotecnico":    "Nemotécnico",
        "fec_vcto":       "Vencimiento",
        "cv_monto":       "CV Monto",
        "cv_num_opes":    "# Oper.",
        "cv_tasa_min":    "Tasa Mín",
        "cv_tasa_max":    "Tasa Máx",
        "cv_tasa_cierre": "Tasa Cierre",
    }).reset_index(drop=True)
    bond_display["Monto Total"] = bond_display["Monto Total"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
    bond_display["CV Monto"]    = bond_display["CV Monto"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    # ── Buyers / Sellers ──────────────────────────────────────────────────────
    def _find_ext_debt_col(df: pd.DataFrame) -> str:
        if _EXT_DEBT_INSTR in df.columns:
            return _EXT_DEBT_INSTR
        for col in df.columns:
            if "external" in col.lower() and "dollar" in col.lower():
                return col
        for col in df.columns:
            if col not in ("sector", "total"):
                return col
        raise KeyError(f"Cannot find external-debt column in {list(df.columns)}")

    def _side(key, col_label):
        df  = pti_data[key]
        col = _find_ext_debt_col(df)
        out = (
            df[["sector", col]]
            .rename(columns={"sector": "Sector", col: col_label})
            .sort_values(col_label, ascending=False)
            .reset_index(drop=True)
        )
        out[col_label] = out[col_label].round().astype("Int64")
        return out

    buyers  = _side("buyers",  "Monto Comprado")
    sellers = _side("sellers", "Monto Vendido")
    total   = df_bonds["total_monto"].sum()

    report_bytes = generate_to_bytes(xlsx_bytes, report_date).getvalue()
    return bond_display, buyers, sellers, total, report_bytes


@st.cache_data(show_spinner=False)
def _load_from_file(xlsx_path: str, report_date: str):
    """Load dashboard data from a saved history file (no network call)."""
    return _parse(Path(xlsx_path).read_bytes(), report_date)


@st.cache_data(show_spinner=False)
def _load(url: str, report_date: str):
    """Download xlsx for the given report URL, parse it, and build all dashboard data."""
    return _parse(download_xlsx_bytes(url), report_date)



# ── Build date catalogue: history files + live API ────────────────────────────
# History files saved by GitHub Actions (data/history/*.xlsx)
def _date_from_xlsx(p: Path) -> str:
    parts = p.stem.split("_")
    try:
        return f"{parts[1]}-{parts[2]}-{parts[3]}"
    except IndexError:
        return p.stem

history_files = {
    _date_from_xlsx(p): str(p)
    for p in sorted(_HISTORY_DIR.glob("BoletinDiario_*.xlsx"))
    if len(p.stem.split("_")) >= 4
}

# Live API dates (re-checked every hour)
with st.spinner("Consultando BVC..."):
    try:
        reports    = _get_reports()
        api_map    = {r["date"]: r["attached"]["url"] for r in reports}
    except Exception as e:
        st.warning(f"No se pudo conectar con la API de BVC: {e}")
        api_map = {}

# Merge: history takes priority; API fills in recent dates not yet saved
all_dates   = sorted(set(history_files) | set(api_map), reverse=True)
today       = date.today().isoformat()

if not all_dates:
    st.error("No hay datos disponibles. Revisa la conexión o el repositorio.")
    st.stop()

# ── Sidebar ────────────────────────────────────────────────────────────────────
chosen_date = st.sidebar.selectbox("📅 Fecha del informe", all_dates)

if today not in all_dates:
    st.sidebar.info(
        f"ℹ️ El boletín de hoy ({today}) aún no está disponible.  \n"
        "Mostrando el más reciente."
    )

st.sidebar.markdown("---")
st.sidebar.markdown(f"**{len(all_dates)} fechas disponibles**")
st.sidebar.markdown("BVC · RF-Mercado Secundario  \nRF-PorTipoInver")

# ── Load & parse ───────────────────────────────────────────────────────────────
with st.spinner("Cargando datos..."):
    try:
        if chosen_date in history_files:
            bond_df, buyers_df, sellers_df, grand_total, report_bytes = _load_from_file(
                history_files[chosen_date], chosen_date
            )
        else:
            bond_df, buyers_df, sellers_df, grand_total, report_bytes = _load(
                api_map[chosen_date], chosen_date
            )
    except Exception as e:
        st.error(f"Error procesando el informe: {e}")
        st.stop()

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 Bonos Deuda Pública Externa USD")
st.caption(f"Informe diario · {chosen_date} · Sistema de Registro (OTC)")

if grand_total == 0:
    st.warning(
        f"No hubo negociaciones de Bonos Deuda Pública Externa USD el **{chosen_date}**. "
        "Selecciona otra fecha en el panel izquierdo."
    )
    st.stop()

col_metric, col_download = st.columns([3, 1])
with col_metric:
    st.metric(label="Monto Total Negociado", value=f"${grand_total:,.0f} M")
with col_download:
    st.download_button(
        label="⬇️ Descargar Informe Excel",
        data=report_bytes,
        file_name=f"RF_Report_{chosen_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")

# ── Section 1: Bond detail ─────────────────────────────────────────────────────
st.subheader("Bonos BGLT")

st.dataframe(
    bond_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Monto Total":    st.column_config.TextColumn("Monto Total"),
        "CV Monto":       st.column_config.TextColumn("CV Monto"),
        "# Oper.":        st.column_config.NumberColumn("# Oper.",      format="%d"),
        "Tasa Mín":       st.column_config.NumberColumn("Tasa Mín",     format="%.4f"),
        "Tasa Máx":       st.column_config.NumberColumn("Tasa Máx",     format="%.4f"),
        "Tasa Cierre":    st.column_config.NumberColumn("Tasa Cierre",  format="%.4f"),
        "Vencimiento":    st.column_config.TextColumn("Vencimiento"),
    },
)

st.markdown("---")

# ── Section 2: Buyers / Sellers side by side ──────────────────────────────────
st.subheader("Distribución por Tipo de Inversionista")

col_b, col_s = st.columns(2)

def _render_side(col, df: pd.DataFrame, monto_col: str, title: str):
    with col:
        st.markdown(f"#### {title}")
        active = df[df[monto_col] > 0]
        zero   = df[df[monto_col] == 0]

        display = active[[monto_col, "Sector"]].copy()
        display[monto_col] = display[monto_col].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "")
        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
            column_config={
                monto_col: st.column_config.TextColumn(monto_col),
            },
        )
        st.caption(
            f"**Total: ${active[monto_col].sum():,.0f} M** · "
            f"{len(active)} sector(es) activo(s) · "
            f"{len(zero)} sin operaciones"
        )

        if not zero.empty:
            with st.expander(f"Sectores sin operaciones ({len(zero)})"):
                st.dataframe(zero[["Sector"]], use_container_width=True, hide_index=True)

_render_side(col_b, buyers_df,  "Monto Comprado", "🟢 Compradores")
_render_side(col_s, sellers_df, "Monto Vendido",  "🔴 Vendedores")
