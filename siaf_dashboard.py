import re
import io
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="SIAF Dashboard - Peru Compras", layout="wide")
st.title("SIAF Dashboard - Peru Compras")

st.markdown(
    "Carga el **Excel SIAF** y obtén resúmenes de **PIA, PIM, Certificado, Comprometido, Devengado, Saldo PIM y % Avance**. "
    "Lee por defecto **A:CH** (base hasta CI). "
    "Construye el **clasificador concatenado** y lo **normaliza para que siempre comience con `2.`**; además agrega una **descripción jerárquica**. "
    "Incluye filtros, pivotes, serie mensual y descarga a Excel."
)

# =========================
# Sidebar / parámetros
# =========================
with st.sidebar:
    st.header("Parámetros de lectura")
    uploaded = st.file_uploader("Archivo SIAF (.xlsx)", type=["xlsx"])
    usecols = st.text_input(
        "Rango de columnas (Excel)",
        "A:CH",
        help="Lectura fija para asegurar columnas CI–EC",
        disabled=True,
    )
    sheet_name = st.text_input("Nombre de hoja (opcional)", "", help="Déjalo vacío para autodetección.")
    header_row_excel = st.number_input("Fila de encabezados (Excel, 1=primera)", min_value=1, value=4)
    detect_header = st.checkbox("Autodetectar encabezado", value=True)
    st.markdown("---")
    st.header("Reglas CI–EC")
    current_month = st.number_input("Mes actual (1-12)", min_value=1, max_value=12, value=9)
    riesgo_umbral = st.number_input("Umbral de avance mínimo (%)", min_value=0, max_value=100, value=60)
    meta_avance = st.number_input("Meta de avance al cierre (%)", min_value=0, max_value=100, value=95)
    st.caption("Se marca riesgo_devolucion si Avance% < Umbral.")

# Mapeo de códigos de sec_func a nombres
SEC_FUNC_MAP = {
    1: "PI_2",
    2: "DCEME",
    3: "DE",
    4: "PI_1",
    5: "OPP",
    6: "JEF",
    7: "GG",
    8: "OAUGD",
    9: "OTI",
    10: "OA",
    11: "OC",
    12: "OAJ",
    13: "RRHH",
    14: "OCI",
    15: "DCEME15",
    16: "DETN16",
    21: "DETN21",
    22: "DETN22",
}
SEC_FUNC_MAP.update({str(k): v for k, v in SEC_FUNC_MAP.items()})

_sec_func_pattern = re.compile(r"^\s*0*(\d+)")


def map_sec_func(value):
    """Normaliza y reemplaza los códigos *sec_func* por sus áreas."""
    if pd.isna(value):
        return value

    if isinstance(value, (int, np.integer)):
        key = int(value)
        return SEC_FUNC_MAP.get(key, SEC_FUNC_MAP.get(str(key), value))

    if isinstance(value, float) and value.is_integer():
        key = int(value)
        return SEC_FUNC_MAP.get(key, SEC_FUNC_MAP.get(str(key), value))

    text = str(value).strip()
    if not text:
        return value

    match = _sec_func_pattern.match(text)
    if match:
        key_str = match.group(1)
        key_int = int(key_str)
        mapped = SEC_FUNC_MAP.get(key_int, SEC_FUNC_MAP.get(key_str))
        if mapped is not None:
            return mapped

    return SEC_FUNC_MAP.get(text, value)


AMOUNT_KEYWORDS = (
    "mto",
    "devengado",
    "saldo",
    "pia",
    "pim",
    "certificado",
    "compro",
    "monto",
    "actual",
    "real",
    "necesario",
    "estimado",
    "proyeccion",
)
EXCLUDE_ROUND_COLS = {"mes", "rank_acum", "rank_mes", "n"}
Z_SCORE_95 = 1.96


def _format_amount(value):
    return "" if pd.isna(value) else f"{value:,.2f}"


def _format_percent(value):
    return "" if pd.isna(value) else f"{value:.2f}%"


def round_numeric_for_reporting(df):
    """Round monetary/percentage numeric columns to two decimals without altering counts."""
    df = df.copy()
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        if col in EXCLUDE_ROUND_COLS:
            continue
        lower = col.lower()
        if col.endswith("%"):
            df[col] = df[col].round(2)
        elif any(keyword in lower for keyword in AMOUNT_KEYWORDS):
            df[col] = df[col].round(2)
    return df


def build_style_formatters(df):
    """Return formatter dict for Streamlit Styler with 2-decimal monetary and percent columns."""
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    formatters = {}
    for col in numeric_cols:
        if col in EXCLUDE_ROUND_COLS:
            continue
        lower = col.lower()
        if col.endswith("%"):
            formatters[col] = _format_percent
        elif any(keyword in lower for keyword in AMOUNT_KEYWORDS):
            formatters[col] = _format_amount
    return formatters


def join_unique_nonempty(values, sep="\n"):
    """Join unique, non-empty string representations preserving order."""

    seen = []
    for value in values:
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not text or text.lower() == "nan":
            continue
        if text not in seen:
            seen.append(text)
    return sep.join(seen)


def compose_email_body(template, row, meta_avance):
    """Format the user-provided email template with area metrics."""

    def _safe_float(value):
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    context = {
        "area": row.get("sec_func", ""),
        "avance_acum": _safe_float(row.get("avance_acum_%", 0.0)),
        "avance_mes": _safe_float(row.get("avance_mes_%", 0.0)),
        "pim": _safe_float(row.get("mto_pim", 0.0)),
        "devengado": _safe_float(row.get("devengado", 0.0)),
        "devengado_mes": _safe_float(row.get("devengado_mes", 0.0)),
        "meta": _safe_float(meta_avance),
    }
    return template.format(**context)

# =========================
# Utilitarios de carga
# =========================
def autodetect_sheet_and_header(xls, excel_bytes, usecols, user_sheet, header_guess):
    """
    Busca la hoja y la fila que luce como encabezado (contenga 'ano_eje', 'pim', 'pia', 'mto_', 'devenga', 'girado').
    Retorna (sheet_name, header_row_index_pandas).
    """
    candidate_sheets = [user_sheet] if user_sheet else xls.sheet_names
    for s in candidate_sheets:
        try:
            tmp = pd.read_excel(excel_bytes, sheet_name=s, header=None, usecols=usecols, nrows=12)
        except Exception:
            continue
        for r in range(min(8, len(tmp))):
            row_vals = tmp.iloc[r].astype(str).str.lower().tolist()
            hits = sum(int(any(k in v for k in ["ano_eje", "pim", "pia", "mto_", "devenga", "girado"])) for v in row_vals)
            if hits >= 2:
                return s, r
    # Fallback: primera hoja y fila indicada por el usuario - 1 (a índice 0)
    return xls.sheet_names[0], header_guess - 1


def load_data(excel_bytes, usecols, sheet_name, header_row_excel, autodetect=True):
    xls = pd.ExcelFile(excel_bytes)
    if autodetect:
        s, hdr = autodetect_sheet_and_header(xls, excel_bytes, usecols, sheet_name, header_row_excel)
        df = pd.read_excel(excel_bytes, sheet_name=s, header=hdr, usecols=usecols)
    else:
        hdr = header_row_excel - 1
        s = sheet_name if sheet_name else xls.sheet_names[0]
        df = pd.read_excel(excel_bytes, sheet_name=s, header=hdr, usecols=usecols)

    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df, s

# =========================
# Cálculos CI–EC
# =========================
def find_monthly_columns(df, prefix):
    return [f"{prefix}{i:02d}" for i in range(1, 13) if f"{prefix}{i:02d}" in df.columns]


def ensure_ci_ec_steps(df, month, umbral):
    """
    Crea/asegura columnas claves si no existen:
    - devengado (suma mto_devenga_01..12)
    - devengado_mes (columna del mes seleccionado)
    - saldo_pim (pim - devengado)
    - avance_% (devengado/pim)
    - riesgo_devolucion (avance_% < umbral)
    - area (vacía si no existe)
    """
    df = df.copy()
    dev_cols = find_monthly_columns(df, "mto_devenga_")

    if "devengado" not in df.columns:
        df["devengado"] = df[dev_cols].sum(axis=1) if dev_cols else 0.0

    col_mes = f"mto_devenga_{int(month):02d}"
    if "devengado_mes" not in df.columns:
        df["devengado_mes"] = df[col_mes] if col_mes in df.columns else 0.0

    if "saldo_pim" not in df.columns:
        df["saldo_pim"] = np.where(df.get("mto_pim", 0) > 0, df["mto_pim"] - df["devengado"], 0.0)

    if "avance_%" not in df.columns:
        df["avance_%"] = np.where(df.get("mto_pim", 0) > 0, df["devengado"] / df["mto_pim"] * 100.0, 0.0)

    if "riesgo_devolucion" not in df.columns:
        df["riesgo_devolucion"] = df["avance_%"] < float(umbral)

    if "area" not in df.columns:
        df["area"] = ""

    return df

# =========================
# Clasificador concatenado
# =========================
_code_re = re.compile(r"^\s*(\d+(?:\.\d+)*)")


def extract_code(text):
    """Extrae el prefijo numérico (con puntos) de un texto tipo '2.1.1 Bienes y servicios'."""
    if pd.isna(text):
        return ""
    s = str(text).strip()
    m = _code_re.match(s)
    return m.group(1) if m else ""


def last_segment(code):
    return code.split(".")[-1] if code else ""


def concat_hierarchy(gen, sub, subdet, esp, espdet):
    """
    Concatena jerárquicamente evitando duplicados:
    generica.subgenerica.subgenerica_det.especifica.especifica_det
    """
    parts = []
    if gen:
        parts.append(gen)
    for child in [sub, subdet, esp, espdet]:
        if not child:
            continue
        if parts and (child.startswith(parts[-1] + ".") or child.startswith(parts[0] + ".")):
            parts.append(child)
        else:
            if parts:
                parts.append(parts[-1] + "." + last_segment(child))
            else:
                parts.append(child)
    return parts[-1] if parts else ""


def normalize_clasificador(code):
    """
    Regla: todo clasificador debe comenzar con '2.'.
    - Si está vacío => '2.'
    - Si no inicia con '2.' => anteponer '2.'
    """
    if not code:
        return "2."
    return code if code.startswith("2.") else "2." + code


def desc_only(text):
    """Devuelve solo la descripción (lo que va después del primer punto)."""
    if pd.isna(text):
        return ""
    s = str(text)
    return s.split(".", 1)[1].strip() if "." in s else s


def build_classifier_columns(df):
    """
    Crea columnas:
    - gen_cod, sub_cod, subdet_cod, esp_cod, espdet_cod (códigos numéricos)
    - clasificador_cod (concatenado y normalizado con 2.)
    - generica_desc, subgenerica_desc, subgenerica_det_desc, especifica_desc, especifica_det_desc
    - clasificador_desc (descripción jerárquica)
    """
    df = df.copy()
    gen = df.get("generica", "")
    sub = df.get("subgenerica", "")
    subdet = df.get("subgenerica_det", "")
    esp = df.get("especifica", "")
    espdet = df.get("especifica_det", "")

    df["gen_cod"] = gen.map(extract_code) if "generica" in df.columns else ""
    df["sub_cod"] = sub.map(extract_code) if "subgenerica" in df.columns else ""
    df["subdet_cod"] = subdet.map(extract_code) if "subgenerica_det" in df.columns else ""
    df["esp_cod"] = esp.map(extract_code) if "especifica" in df.columns else ""
    df["espdet_cod"] = espdet.map(extract_code) if "especifica_det" in df.columns else ""

    df["clasificador_cod"] = [
        normalize_clasificador(concat_hierarchy(g, s, sd, e, ed))
        for g, s, sd, e, ed in zip(
            df["gen_cod"], df["sub_cod"], df["subdet_cod"], df["esp_cod"], df["espdet_cod"]
        )
    ]

    df["generica_desc"] = gen.map(desc_only) if "generica" in df.columns else ""
    df["subgenerica_desc"] = sub.map(desc_only) if "subgenerica" in df.columns else ""
    df["subgenerica_det_desc"] = subdet.map(desc_only) if "subgenerica_det" in df.columns else ""
    df["especifica_desc"] = esp.map(desc_only) if "especifica" in df.columns else ""
    df["especifica_det_desc"] = espdet.map(desc_only) if "especifica_det" in df.columns else ""

    df["clasificador_desc"] = (
        df["generica_desc"].fillna("")
        + " > " + df["subgenerica_desc"].fillna("")
        + " > " + df["subgenerica_det_desc"].fillna("")
        + " > " + df["especifica_desc"].fillna("")
        + " > " + df["especifica_det_desc"].fillna("")
    ).str.strip(" >")

    return df

# =========================
# Pivote / resumen por grupo
# =========================
def pivot_exec(df, group_col, dev_cols):
    cols = []
    if "mto_pia" in df.columns:
        cols.append("mto_pia")
    if "mto_pim" in df.columns:
        cols.append("mto_pim")
    if "mto_certificado" in df.columns:
        cols.append("mto_certificado")
    if "mto_compro_anual" in df.columns:
        cols.append("mto_compro_anual")
    if dev_cols:
        cols.append("devengado")

    if "devengado" not in df.columns and dev_cols:
        df = df.copy()
        df["devengado"] = df[dev_cols].sum(axis=1)

    g = df.groupby(group_col, dropna=False)[cols].sum().reset_index()

    if "mto_pim" in g.columns and "devengado" in g.columns:
        g["saldo_pim"] = g["mto_pim"] - g["devengado"]
        g["avance_%"] = np.where(g["mto_pim"] > 0, g["devengado"] / g["mto_pim"] * 100.0, 0.0)

    return g


def to_excel_download(
    resumen,
    avance,
    proyeccion=None,
    ritmo=None,
    leaderboard=None,
    reporte_siaf=None,
):
    wb = Workbook()
    wb.remove(wb.active)

    def add_table_with_chart(df, sheet_name):
        ws = wb.create_sheet(sheet_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        if ws.max_row <= 1:
            return
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tbl = Table(displayName=f"Tbl{sheet_name[:20].replace(' ','_')}", ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tbl)

        num_cols = [i + 2 for i, c in enumerate(df.columns[1:]) if pd.api.types.is_numeric_dtype(df[c])]
        if num_cols:
            chart = BarChart()
            data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row, max_col=max(num_cols))
            cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.title = sheet_name
            chart.height = 7
            chart.width = 15
            ws.add_chart(chart, f"{get_column_letter(ws.max_column + 2)}2")

    add_table_with_chart(resumen, "Resumen")
    add_table_with_chart(avance, "Avance")
    if proyeccion is not None and not proyeccion.empty:
        add_table_with_chart(proyeccion, "Proyeccion")
    if ritmo is not None and not ritmo.empty:
        add_table_with_chart(ritmo, "Ritmo")
    if leaderboard is not None and not leaderboard.empty:
        add_table_with_chart(leaderboard, "Leaderboard")
    if reporte_siaf is not None and not reporte_siaf.empty:
        add_table_with_chart(reporte_siaf, "Reporte_SIAF")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =========================
# Carga del archivo
# =========================
if uploaded is None:
    st.info("Sube tu archivo Excel SIAF para empezar. Usa A:CH para incluir pasos CI–EC.")
    st.stop()

try:
    df, used_sheet = load_data(uploaded, usecols, sheet_name.strip() or None, int(header_row_excel), autodetect=detect_header)
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

st.success(f"Leída la hoja '{used_sheet}' con {df.shape[0]} filas y {df.shape[1]} columnas.")

if "sec_func" in df.columns:
    df["sec_func"] = df["sec_func"].apply(map_sec_func)

# =========================
# Filtros
# =========================
st.subheader("Filtros")
filter_cols = [c for c in df.columns if any(k in c for k in [
    "unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func",
    "departamento_meta","provincia_meta","distrito_meta","area"
])]

cols_f = st.columns(3)
selected_filters = {}
for i, c in enumerate(filter_cols):
    with cols_f[i % 3]:
        vals = sorted([str(x) for x in df[c].dropna().unique().tolist()])
        if len(vals) > 1:
            pick = st.multiselect(c, options=vals, default=[])
            if pick:
                selected_filters[c] = set(pick)

mask = pd.Series(True, index=df.index)
for c, allowed in selected_filters.items():
    mask &= df[c].astype(str).isin(allowed)
df_f = df[mask].copy()

# =========================
# Aplicar CI–EC + Clasificador
# =========================
df_proc = ensure_ci_ec_steps(df_f, current_month, riesgo_umbral)
df_proc = build_classifier_columns(df_proc)

# Totales globales para el resumen ejecutivo
_tot_pia = float(df_proc.get("mto_pia", 0).sum())
_tot_pim = float(df_proc.get("mto_pim", 0).sum())
_tot_dev = float(df_proc.get("devengado", 0).sum())
_tot_cert = float(df_proc.get("mto_certificado", 0).sum()) if "mto_certificado" in df_proc.columns else 0.0
_tot_comp = float(df_proc.get("mto_compro_anual", 0).sum()) if "mto_compro_anual" in df_proc.columns else 0.0
_saldo_pim = _tot_pim - _tot_dev if _tot_pim else 0.0
_avance_global = (_tot_dev / _tot_pim * 100.0) if _tot_pim else 0.0

dev_cols = [c for c in df_proc.columns if c.startswith("mto_devenga_")]

_group_options = [c for c in df_proc.columns if c in [
    "clasificador_cod","unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func","area"
]]
_group_default = _group_options.index("clasificador_cod") if "clasificador_cod" in _group_options else 0
_group_col = st.selectbox("Agrupar por", options=_group_options, index=_group_default)
_group_values = ["(Todos)"] + sorted(df_proc[_group_col].dropna().astype(str).unique().tolist())
_group_filter = st.selectbox(f"Filtrar {_group_col}", options=_group_values, index=0)

df_view = df_proc if _group_filter == "(Todos)" else df_proc[df_proc[_group_col].astype(str) == _group_filter]

# Datos precomputados para cada apartado
pivot = pivot_exec(df_view, _group_col, dev_cols)

_ci_cols = [
    "clasificador_cod", "clasificador_desc",
    "generica","subgenerica","subgenerica_det","especifica","especifica_det",
    "mto_pia","mto_pim","mto_certificado","mto_compro_anual",
    "devengado_mes","devengado","saldo_pim","avance_%","riesgo_devolucion"
]
_ci_cols = [c for c in _ci_cols if c in df_view.columns]
df_ci = df_view[_ci_cols].head(300) if _ci_cols else pd.DataFrame()

_consol_cols = [
    c
    for c in [
        "mto_pia",
        "mto_pim",
        "mto_certificado",
        "mto_compro_anual",
        "devengado_mes",
        "devengado",
        "saldo_pim",
    ]
    if c in df_view.columns
]
consolidado = pd.DataFrame()
if _consol_cols:
    consolidado = df_view.groupby(
        ["clasificador_cod","clasificador_desc","generica","subgenerica","subgenerica_det","especifica","especifica_det"],
        dropna=False
    )[_consol_cols].sum().reset_index()
    if "mto_pim" in consolidado.columns and "devengado" in consolidado.columns:
        consolidado["avance_%"] = np.where(
            consolidado["mto_pim"] > 0,
            consolidado["devengado"] / consolidado["mto_pim"] * 100.0,
            0.0,
        )

avance_series = pd.DataFrame()
if dev_cols and "mto_pim" in df_view.columns:
    month_map = {f"mto_devenga_{i:02d}": i for i in range(1, 13)}
    dev_series = df_view[dev_cols].sum().reset_index()
    dev_series.columns = ["col", "devengado"]
    dev_series["mes"] = dev_series["col"].map(month_map)
    dev_series = dev_series.dropna(subset=["mes"]).sort_values("mes")
    pim_total = df_view["mto_pim"].sum()
    dev_series["devengado"] = dev_series["devengado"].fillna(0.0)
    dev_series["acumulado"] = dev_series["devengado"].cumsum()
    dev_series["%_acumulado"] = np.where(
        pim_total > 0,
        dev_series["acumulado"] / pim_total * 100.0,
        0.0,
    )
    dev_series["%_acumulado"] = dev_series["%_acumulado"].round(2)
    avance_series = dev_series[["mes", "devengado", "%_acumulado"]]

ritmo_df = pd.DataFrame()
leaderboard_df = pd.DataFrame()
alert_df = pd.DataFrame()
reporte_siaf_df = pd.DataFrame()
proyeccion_wide = pd.DataFrame()

# Navegación por apartados
(
    tab_resumen,
    tab_agrup,
    tab_ci,
    tab_consol,
    tab_avance,
    tab_gestion,
    tab_reporte,
    tab_descarga,
) = st.tabs([
    "Resumen ejecutivo",
    "Agrupaciones",
    "Procesos CI–EC",
    "Consolidado",
    "Avance mensual",
    "Ritmo y alertas",
    "Reporte SIAF",
    "Descargas",
])

with tab_resumen:
    st.header("Resumen ejecutivo (totales)")
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric("PIA", f"S/ {_tot_pia:,.2f}")
    k2.metric("PIM", f"S/ {_tot_pim:,.2f}")
    k3.metric("Certificado", f"S/ {_tot_cert:,.2f}")
    k4.metric("Comprometido", f"S/ {_tot_comp:,.2f}")
    k5.metric("Devengado (YTD)", f"S/ {_tot_dev:,.2f}")
    k6.metric("Saldo PIM", f"S/ {_saldo_pim:,.2f}")
    k7.metric("Avance", f"{_avance_global:.2f}%")

with tab_agrup:
    st.header("Vistas por agrupación")
    pivot_display = round_numeric_for_reporting(pivot)
    fmt_pivot = build_style_formatters(pivot_display)
    pivot_style = pivot_display.style
    if "avance_%" in pivot_display.columns:
        pivot_style = pivot_style.applymap(
            lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
            subset=["avance_%"],
        )
    if fmt_pivot:
        pivot_style = pivot_style.format(fmt_pivot)
    st.dataframe(pivot_style, use_container_width=True)

with tab_ci:
    st.header("Procesos CI–EC (monto vinculado a su cadena)")
    if df_ci.empty:
        st.info("No hay datos disponibles para esta vista.")
    else:
        df_ci_display = round_numeric_for_reporting(df_ci)
        fmt_ci = build_style_formatters(df_ci_display)
        ci_style = df_ci_display.style
        if "avance_%" in df_ci_display.columns:
            ci_style = ci_style.applymap(
                lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
                subset=["avance_%"],
            )
        if fmt_ci:
            ci_style = ci_style.format(fmt_ci)
        st.dataframe(ci_style, use_container_width=True)

with tab_consol:
    st.header("Consolidado por clasificador")
    if consolidado.empty:
        st.info("No hay información consolidada para mostrar.")
    else:
        consol_display = round_numeric_for_reporting(consolidado.head(500))
        fmt_consol = build_style_formatters(consol_display)
        consol_style = consol_display.style
        if "avance_%" in consol_display.columns:
            consol_style = consol_style.applymap(
                lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
                subset=["avance_%"],
            )
        if fmt_consol:
            consol_style = consol_style.format(fmt_consol)
        st.dataframe(consol_style, use_container_width=True)

with tab_avance:
    st.header("Avance mensual interactivo")
    if avance_series.empty:
        st.info("No hay información de devengado mensual para graficar.")
    else:
        avance_display = avance_series.copy()
        bar = (
            alt.Chart(avance_display)
            .mark_bar(color="#1f77b4")
            .encode(
                x=alt.X("mes:O", title="Mes"),
                y=alt.Y("devengado:Q", title="Devengado", axis=alt.Axis(format=",.2f")),
                tooltip=[
                    alt.Tooltip("mes", title="Mes"),
                    alt.Tooltip("devengado", title="Devengado", format=",.2f"),
                    alt.Tooltip("%_acumulado", title="% acumulado", format=".2f"),
                ],
            )
        )
        line = (
            alt.Chart(avance_display)
            .mark_line(color="#ff7f0e", point=True)
            .encode(
                x=alt.X("mes:O", title="Mes"),
                y=alt.Y("%_acumulado:Q", title="% acumulado", axis=alt.Axis(format=".2f")),
                tooltip=[
                    alt.Tooltip("mes", title="Mes"),
                    alt.Tooltip("%_acumulado", title="% acumulado", format=".2f"),
                ],
            )
        )
        chart = alt.layer(bar, line).resolve_scale(y="independent").properties(width=520, height=260)
        st.altair_chart(chart, use_container_width=False)

        avance_table = round_numeric_for_reporting(avance_display)
        fmt_avance = build_style_formatters(avance_table)
        avance_style = avance_table.style
        if "%_acumulado" in avance_table.columns:
            avance_style = avance_style.applymap(
                lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
                subset=["%_acumulado"],
            )
        if fmt_avance:
            avance_style = avance_style.format(fmt_avance)
        st.dataframe(avance_style, use_container_width=True)

with tab_gestion:
    st.header("Ritmo requerido por proceso")
    if "mto_pim" not in df_view.columns:
        st.info("Se requiere la columna mto_pim para calcular el ritmo requerido.")
    else:
        remaining_months = max(12 - current_month, 1)
        pim_total = df_view["mto_pim"].sum()
        processes = []
        for col, label in [("mto_certificado", "Certificar"), ("mto_compro_anual", "Comprometer"), ("devengado", "Devengar")]:
            total = df_view.get(col, pd.Series(dtype=float)).sum()
            actual_avg = total / max(current_month, 1)
            needed = max(pim_total - total, 0)
            required_avg = needed / remaining_months
            processes.append({"Proceso": label, "Actual": actual_avg, "Necesario": required_avg})
        ritmo_df = pd.DataFrame(processes)
        ritmo_df = round_numeric_for_reporting(ritmo_df)
        ritmo_melt = ritmo_df.melt("Proceso", var_name="Tipo", value_name="Monto")
        chart_ritmo = (
            alt.Chart(ritmo_melt)
            .mark_bar()
            .encode(
                x=alt.X("Proceso:N"),
                y=alt.Y("Monto:Q", axis=alt.Axis(format=",.2f")),
                color="Tipo:N",
                tooltip=["Proceso", "Tipo", alt.Tooltip("Monto", format=",.2f")],
            )
            .properties(width=600, height=300)
        )
        st.altair_chart(chart_ritmo, use_container_width=False)

    st.header("Top áreas con menor avance")
    if "sec_func" in df_view.columns and "mto_pim" in df_view.columns:
        agg_cols = ["mto_pim", "devengado", "devengado_mes"]
        if "mto_certificado" in df_view.columns:
            agg_cols.insert(1, "mto_certificado")
        agg_sec = df_view.groupby("sec_func", dropna=False)[agg_cols].sum().reset_index()
        if not agg_sec.empty:
            agg_sec["avance_acum_%"] = np.where(agg_sec["mto_pim"] > 0, agg_sec["devengado"] / agg_sec["mto_pim"] * 100.0, 0.0)
            agg_sec["avance_mes_%"] = np.where(
                agg_sec["mto_pim"] > 0, agg_sec["devengado_mes"] / agg_sec["mto_pim"] * 100.0, 0.0
            )
            agg_sec["rank_acum"] = agg_sec["avance_acum_%"].rank(method="dense", ascending=True).astype(int)
            agg_sec["rank_mes"] = agg_sec["avance_mes_%"].rank(method="dense", ascending=True).astype(int)

            max_top = int(agg_sec.shape[0])
            top_default = 5 if max_top >= 5 else max_top
            top_n = st.slider("Número de áreas a mostrar", min_value=1, max_value=max_top, value=top_default)

            leaderboard_df = (
                agg_sec.sort_values(["avance_acum_%", "avance_mes_%"], ascending=[True, True])
                .head(top_n)
                .copy()
            )
            display_cols = [
                "rank_acum",
                "rank_mes",
                "sec_func",
                "mto_pim",
            ]
            if "mto_certificado" in agg_sec.columns:
                display_cols.append("mto_certificado")
            display_cols.extend([
                "devengado",
                "avance_acum_%",
                "devengado_mes",
                "avance_mes_%",
            ])
            leaderboard_df = leaderboard_df[display_cols]

            leaderboard_display = round_numeric_for_reporting(leaderboard_df)
            fmt_leader = build_style_formatters(leaderboard_display)
            highlight = lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else ""
            leader_style = leaderboard_display.style.applymap(
                highlight,
                subset=["avance_acum_%", "avance_mes_%"],
            )
            if fmt_leader:
                leader_style = leader_style.format(fmt_leader)
            st.dataframe(leader_style, use_container_width=True)
    else:
        st.info("Se requieren las columnas sec_func y mto_pim para construir el ranking.")

    st.header("Automatización de alertas por correo")
    if "alert_contacts" not in st.session_state:
        st.session_state["alert_contacts"] = {}
    if "email_subject" not in st.session_state:
        st.session_state["email_subject"] = f"Alerta de avance presupuestal - Mes {int(current_month):02d}"
    if "email_body_template" not in st.session_state:
        st.session_state["email_body_template"] = (
            "Estimado equipo {area},\n\n"
            "El avance acumulado registra {avance_acum:.2f}% y el avance del mes es {avance_mes:.2f}%.\n"
            "PIM: S/ {pim:,.2f}\n"
            "Devengado acumulado: S/ {devengado:,.2f}\n"
            "Devengado del mes: S/ {devengado_mes:,.2f}\n\n"
            "La meta institucional vigente es {meta:.0f}%. Por favor revisen las acciones necesarias para mejorar la ejecución.\n\n"
            "Saludos,\n"
            "Equipo de Presupuesto"
        )

    if not leaderboard_df.empty:
        mask_acum = leaderboard_df.get("avance_acum_%", pd.Series(dtype=float)) < float(riesgo_umbral)
        mask_mes = leaderboard_df.get("avance_mes_%", pd.Series(dtype=float)) < float(riesgo_umbral)
        risk_mask = (mask_acum.fillna(False)) | (mask_mes.fillna(False))
        if risk_mask.any():
            alert_df = leaderboard_df.loc[risk_mask].copy()

    if alert_df.empty:
        st.info("No hay áreas con avance por debajo del umbral definido. Ajusta los filtros o el umbral para generar alertas.")
    else:
        alert_display = round_numeric_for_reporting(alert_df.copy())
        fmt_alert = build_style_formatters(alert_display)
        highlight_alert = lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else ""
        alert_style = alert_display.style.applymap(
            highlight_alert,
            subset=[c for c in ["avance_acum_%", "avance_mes_%"] if c in alert_display.columns],
        )
        if fmt_alert:
            alert_style = alert_style.format(fmt_alert)
        st.dataframe(alert_style, use_container_width=True)

        alert_areas = sorted(alert_df["sec_func"].astype(str).unique())
        for area in alert_areas:
            st.session_state["alert_contacts"].setdefault(area, "")

        st.markdown("### Contactos por área en riesgo")
        contact_df = pd.DataFrame(
            {
                "Área": alert_areas,
                "Correo": [st.session_state["alert_contacts"].get(area, "") for area in alert_areas],
            }
        )
        edited_contacts = st.data_editor(
            contact_df,
            key="alert_contacts_editor",
            num_rows="fixed",
            use_container_width=True,
        )
        if st.button("Guardar contactos", key="save_contacts"):
            updated_contacts = {}
            for area, email in zip(alert_areas, edited_contacts["Correo"].tolist()):
                if isinstance(email, str):
                    clean_email = email.strip()
                elif pd.notna(email):
                    clean_email = str(email).strip()
                else:
                    clean_email = ""
                if clean_email:
                    updated_contacts[area] = clean_email
            st.session_state["alert_contacts"] = updated_contacts
            st.success("Contactos actualizados correctamente.")

        missing_contacts = [area for area in alert_areas if not st.session_state["alert_contacts"].get(area)]
        if missing_contacts:
            st.info("Faltan correos para: " + ", ".join(missing_contacts))

        st.markdown("### Configurar envío de correos")
        sender_email = st.text_input("Cuenta Outlook (remitente)")
        app_password = st.text_input("Contraseña o app password de Outlook", type="password")
        smtp_server = st.text_input("Servidor SMTP", value="smtp.office365.com")
        smtp_port = st.number_input("Puerto SMTP", min_value=1, max_value=65535, value=587, step=1)

        subject = st.text_input("Asunto del correo", key="email_subject")
        body_template = st.text_area(
            "Plantilla del mensaje (usa llaves para reemplazos como {area}, {avance_acum}, {avance_mes}, {pim})",
            key="email_body_template",
            height=220,
        )

        preview_area = st.selectbox("Vista previa del mensaje", alert_areas, key="preview_area")
        if preview_area:
            preview_row = alert_df[alert_df["sec_func"].astype(str) == preview_area].iloc[0]
            preview_body = compose_email_body(body_template, preview_row, meta_avance)
            st.code(preview_body)

        if st.button("Enviar correos de alerta", key="send_alerts"):
            if not sender_email or not app_password:
                st.error("Debes ingresar la cuenta y la contraseña o app password de Outlook.")
            else:
                active_contacts = {
                    area: email.strip()
                    for area, email in st.session_state["alert_contacts"].items()
                    if isinstance(email, str) and email.strip()
                }
                if not active_contacts:
                    st.warning("No hay correos configurados para las áreas en riesgo.")
                else:
                    messages = []
                    for area, recipient in active_contacts.items():
                        row_match = alert_df[alert_df["sec_func"].astype(str) == area]
                        if row_match.empty:
                            continue
                        row = row_match.iloc[0]
                        body = compose_email_body(body_template, row, meta_avance)
                        msg = EmailMessage()
                        msg["Subject"] = subject
                        msg["From"] = sender_email
                        msg["To"] = recipient
                        msg.set_content(body)
                        messages.append(msg)
                    if not messages:
                        st.warning("No se generaron mensajes para enviar.")
                    else:
                        try:
                            with smtplib.SMTP(smtp_server, int(smtp_port)) as smtp:
                                smtp.starttls()
                                smtp.login(sender_email, app_password)
                                for msg in messages:
                                    smtp.send_message(msg)
                            st.success(f"Se enviaron {len(messages)} alerta(s) correctamente.")
                        except Exception as exc:
                            st.error(f"No se pudieron enviar los correos: {exc}")

with tab_reporte:
    st.header("Reporte SIAF por área, genérica y específica detalle")
    if not all(col in df_view.columns for col in ["sec_func", "generica", "especifica_det"]):
        st.info("Para el reporte SIAF se requieren las columnas sec_func, generica y especifica_det.")
    else:
        siaf_agg_cols = [
            c
            for c in [
                "mto_pia",
                "mto_pim",
                "mto_certificado",
                "mto_compro_anual",
                "devengado_mes",
                "devengado",
                "saldo_pim",
            ]
            if c in df_view.columns
        ]

        if not siaf_agg_cols:
            st.info("No se encontraron columnas monetarias para generar el reporte SIAF por área.")
        else:
            base_group = ["sec_func", "generica", "especifica_det"]
            agg_map = {col: "sum" for col in siaf_agg_cols}
            if "clasificador_cod" in df_view.columns:
                agg_map["clasificador_cod"] = "first"
            if "especifica_det_desc" in df_view.columns:
                agg_map["especifica_det_desc"] = "first"

            reporte_base = (
                df_view.groupby(base_group, dropna=False)
                .agg(agg_map)
                .reset_index()
            )

            if "clasificador_cod" in reporte_base.columns:
                clasificador_cod = reporte_base["clasificador_cod"].fillna("").astype(str).str.strip()
            else:
                clasificador_cod = reporte_base["especifica_det"].map(extract_code).fillna("").astype(str).str.strip()

            if "especifica_det_desc" in reporte_base.columns:
                concepto = reporte_base["especifica_det_desc"].fillna("").astype(str).str.strip()
            else:
                concepto = reporte_base["especifica_det"].map(desc_only).fillna("").astype(str).str.strip()

            tiene_cod = (clasificador_cod != "") & (clasificador_cod != "nan")
            tiene_concepto = (concepto != "") & (concepto != "nan")
            reporte_base["clasificador_cod_concepto"] = np.where(
                tiene_cod & tiene_concepto,
                clasificador_cod + " - " + concepto,
                np.where(tiene_concepto, concepto, clasificador_cod),
            )

            value_sources = {
                "AVANCE DE EJECUCIÓN ACUMULADO": "devengado",
                "PIM": "mto_pim",
                "CERTIFICADO": "mto_certificado",
                "COMPROMETIDO": "mto_compro_anual",
                "DEVENGADO": "devengado_mes",
            }
            for src in value_sources.values():
                if src not in reporte_base.columns:
                    reporte_base[src] = 0.0

            reporte_base = reporte_base[
                reporte_base["clasificador_cod_concepto"].fillna("").astype(str).str.strip() != ""
            ].copy()

            def _label_or_default(value, fallback):
                text = "" if pd.isna(value) else str(value).strip()
                return text if text else fallback

            def _sort_key(label):
                text = _label_or_default(label, "")
                code = extract_code(text)
                if not code:
                    return (tuple(), text)
                parts = []
                for segment in code.split('.'):
                    try:
                        parts.append(int(segment))
                    except ValueError:
                        parts.append(segment)
                return (tuple(parts), text)

            def _format_label(level, text):
                indent = "    " * max(level, 0)
                prefix_map = {0: "", 1: "• ", 2: "- "}
                return f"{indent}{prefix_map.get(level, '- ')}{text}".rstrip()

            records = []
            order_counter = 0

            for sec_value, sec_group in reporte_base.groupby("sec_func", sort=True):
                sec_label = _label_or_default(sec_value, "Sin sec_func")

                def _sum_metric(frame, source):
                    return float(frame[source].fillna(0.0).sum()) if source in frame.columns else 0.0

                sec_metrics = {dest: _sum_metric(sec_group, src) for dest, src in value_sources.items()}
                records.append(
                    {
                        "nivel": 0,
                        "orden": order_counter,
                        "Centro de costo / Genérica de Gasto / Específica de Gasto": _format_label(0, sec_label),
                        **sec_metrics,
                    }
                )
                order_counter += 1

                gen_groups = sorted(
                    sec_group.groupby("generica", dropna=False),
                    key=lambda kv: _sort_key(kv[0]),
                )

                for gen_value, gen_group in gen_groups:
                    gen_label = _label_or_default(gen_value, "Sin genérica")
                    gen_metrics = {dest: _sum_metric(gen_group, src) for dest, src in value_sources.items()}
                    records.append(
                        {
                            "nivel": 1,
                            "orden": order_counter,
                            "Centro de costo / Genérica de Gasto / Específica de Gasto": _format_label(1, gen_label),
                            **gen_metrics,
                        }
                    )
                    order_counter += 1

                    detail_rows = sorted(
                        gen_group.to_dict("records"),
                        key=lambda row: _sort_key(row.get("especifica_det", "")),
                    )
                    for detail_row in detail_rows:
                        spec_label = _label_or_default(detail_row.get("clasificador_cod_concepto", ""), "Sin específica")
                        if not spec_label or spec_label == "Sin específica":
                            continue
                        detail_metrics = {
                            dest: float(detail_row.get(src, 0.0) or 0.0)
                            for dest, src in value_sources.items()
                        }
                        records.append(
                            {
                                "nivel": 2,
                                "orden": order_counter,
                                "Centro de costo / Genérica de Gasto / Específica de Gasto": _format_label(2, spec_label),
                                **detail_metrics,
                            }
                        )
                        order_counter += 1

            if records:
                reporte_siaf_df = pd.DataFrame.from_records(records)
                reporte_siaf_df["% AVANCE DEV /PIM"] = np.where(
                    reporte_siaf_df["PIM"].astype(float) > 0,
                    reporte_siaf_df["AVANCE DE EJECUCIÓN ACUMULADO"].astype(float) / reporte_siaf_df["PIM"].astype(float) * 100.0,
                    0.0,
                )
                reporte_siaf_df = (
                    reporte_siaf_df.sort_values("orden", kind="stable")
                    .drop(columns=["orden", "nivel"], errors="ignore")
                )
                reporte_siaf_df = reporte_siaf_df[[
                    "Centro de costo / Genérica de Gasto / Específica de Gasto",
                    "AVANCE DE EJECUCIÓN ACUMULADO",
                    "PIM",
                    "CERTIFICADO",
                    "COMPROMETIDO",
                    "DEVENGADO",
                    "% AVANCE DEV /PIM",
                ]]
            else:
                reporte_siaf_df = pd.DataFrame(
                    columns=[
                        "Centro de costo / Genérica de Gasto / Específica de Gasto",
                        "AVANCE DE EJECUCIÓN ACUMULADO",
                        "PIM",
                        "CERTIFICADO",
                        "COMPROMETIDO",
                        "DEVENGADO",
                        "% AVANCE DEV /PIM",
                    ]
                )

            reporte_display = round_numeric_for_reporting(reporte_siaf_df)
            fmt_reporte = build_style_formatters(reporte_display)
            reporte_style = reporte_display.style
            if "% AVANCE DEV /PIM" in reporte_display.columns:
                reporte_style = reporte_style.applymap(
                    lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
                    subset=["% AVANCE DEV /PIM"],
                )
            if fmt_reporte:
                reporte_style = reporte_style.format(fmt_reporte)
            st.dataframe(reporte_style, use_container_width=True)

with tab_descarga:
    st.header("Descarga de reportes")
    buf = to_excel_download(
        resumen=round_numeric_for_reporting(pivot.copy()),
        avance=round_numeric_for_reporting(avance_series.copy()),
        proyeccion=proyeccion_wide,
        ritmo=round_numeric_for_reporting(ritmo_df.copy()),
        leaderboard=round_numeric_for_reporting(leaderboard_df.copy()),
        reporte_siaf=round_numeric_for_reporting(reporte_siaf_df.copy()),
    )
    st.download_button(
        "Descargar Excel (Resumen + Avance)",
        data=buf,
        file_name="siaf_resumen_avance.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
