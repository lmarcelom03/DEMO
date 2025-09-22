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
    1: "PI 2",
    2: "DCEME2",
    3: "DE",
    4: "PI 1",
    5: "OPP",
    6: "JEFATURA",
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
    18: "DCEME18",
    19: "DCEME19",
    20: "DETN20",
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
    Busca la hoja y la fila que luce como encabezado (contenga 'ano_eje', 'pim', 'pia', etc.).
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
        # Si el hijo ya trae el prefijo, lo conservamos
        if parts and (child.startswith(parts[-1] + ".") or child.startswith(parts[0] + ".")):
            parts.append(child)
        else:
            # Caso contrario, agregamos solo el último segmento al prefijo anterior
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

    # Descripciones sin código
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

    # Si no existía 'devengado' pero hay columnas mensuales, lo armamos en copia
    if "devengado" not in df.columns and dev_cols:
        df = df.copy()
        df["devengado"] = df[dev_cols].sum(axis=1)

    g = df.groupby(group_col, dropna=False)[cols].sum().reset_index()

    if "mto_pim" in g.columns and "devengado" in g.columns:
        g["saldo_pim"] = g["mto_pim"] - g["devengado"]
        g["avance_%"] = np.where(g["mto_pim"] > 0, g["devengado"] / g["mto_pim"] * 100.0, 0.0)

    return g

def to_excel_download(resumen, avance, proyeccion=None, ritmo=None, leaderboard=None):
    wb = Workbook()
    # remove the default sheet to control ordering
    wb.remove(wb.active)

    def add_table_with_chart(df, sheet_name):
        ws = wb.create_sheet(sheet_name)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        # Skip table/chart creation when there are no data rows
        if ws.max_row <= 1:
            return
        # create an Excel table for easier filtering in the workbook
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tbl = Table(displayName=f"Tbl{sheet_name[:20].replace(' ','_')}", ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tbl)

        # build a bar chart using the first column as categories and remaining numeric columns as data
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

# =========================
# Resumen ejecutivo
# =========================
st.subheader("Resumen ejecutivo (totales)")
dev_cols = [c for c in df_proc.columns if c.startswith("mto_devenga_")]

tot_pia = float(df_proc.get("mto_pia", 0).sum())
tot_pim = float(df_proc.get("mto_pim", 0).sum())
tot_dev = float(df_proc.get("devengado", 0).sum())
tot_cert = float(df_proc.get("mto_certificado", 0).sum()) if "mto_certificado" in df_proc.columns else 0.0
tot_comp = float(df_proc.get("mto_compro_anual", 0).sum()) if "mto_compro_anual" in df_proc.columns else 0.0
saldo_pim = tot_pim - tot_dev if tot_pim else 0.0
avance = (tot_dev / tot_pim * 100.0) if tot_pim else 0.0

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
k1.metric("PIA", f"S/ {tot_pia:,.2f}")
k2.metric("PIM", f"S/ {tot_pim:,.2f}")
k3.metric("Certificado", f"S/ {tot_cert:,.2f}")
k4.metric("Comprometido", f"S/ {tot_comp:,.2f}")
k5.metric("Devengado (YTD)", f"S/ {tot_dev:,.2f}")
k6.metric("Saldo PIM", f"S/ {saldo_pim:,.2f}")
k7.metric("Avance", f"{avance:.2f}%")

# =========================
# Vistas por agrupación
# =========================
st.subheader("Vistas por agrupación")
group_options = [c for c in df_proc.columns if c in [
    "clasificador_cod","unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func","area"
]]
default_idx = group_options.index("clasificador_cod") if "clasificador_cod" in group_options else 0
group_col = st.selectbox("Agrupar por", options=group_options, index=default_idx)

group_vals = ["(Todos)"] + sorted(df_proc[group_col].dropna().astype(str).unique().tolist())
group_val = st.selectbox(f"Filtrar {group_col}", options=group_vals, index=0)
df_view = df_proc if group_val == "(Todos)" else df_proc[df_proc[group_col].astype(str) == group_val]

pivot = pivot_exec(df_view, group_col, dev_cols)
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

# =========================
# Procesos CI–EC (detalle)
# =========================
st.subheader("Procesos CI–EC (monto vinculado a su cadena)")
ci_cols = [
    "clasificador_cod", "clasificador_desc",
    "generica","subgenerica","subgenerica_det","especifica","especifica_det",
    "mto_pia","mto_pim","mto_certificado","mto_compro_anual",
    "devengado_mes","devengado","saldo_pim","avance_%","riesgo_devolucion"
]
ci_cols = [c for c in ci_cols if c in df_view.columns]
df_ci = df_view[ci_cols].head(300)
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

# =========================
# Consolidado por clasificador
# =========================
agg_cols = [
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
consolidado = df_view.groupby(
    ["clasificador_cod","clasificador_desc","generica","subgenerica","subgenerica_det","especifica","especifica_det"],
    dropna=False
)[agg_cols].sum().reset_index()

if "mto_pim" in consolidado.columns and "devengado" in consolidado.columns:
    consolidado["avance_%"] = np.where(consolidado["mto_pim"] > 0, consolidado["devengado"]/consolidado["mto_pim"]*100.0, 0.0)

st.markdown("**Consolidado por clasificador**")
consol_display = consolidado.head(500)
consol_display = round_numeric_for_reporting(consol_display)
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

# =========================
# Serie mensual interactiva
# =========================
avance_series = pd.DataFrame()
proyeccion_wide = pd.DataFrame()
if dev_cols and "mto_pim" in df_view.columns:
    st.subheader("Avance mensual interactivo")
    month_map = {f"mto_devenga_{i:02d}": i for i in range(1, 13)}
    dev_series = df_view[dev_cols].sum().reset_index()
    dev_series.columns = ["col", "monto"]
    dev_series["mes"] = dev_series["col"].map(month_map)
    dev_series = dev_series.sort_values("mes")
    pim_total = df_view["mto_pim"].sum()
    dev_series["contrib_pct"] = np.where(pim_total > 0, dev_series["monto"] / pim_total * 100.0, 0.0)
    dev_series["riesgo"] = dev_series["contrib_pct"] < float(riesgo_umbral)
    dev_series["monto"] = dev_series["monto"].round(2)
    dev_series["contrib_pct"] = dev_series["contrib_pct"].round(2)
    avance_series = dev_series[["mes", "monto", "contrib_pct"]]
    chart = (
        alt.Chart(dev_series)
        .mark_bar()
        .encode(
            x=alt.X("mes:O", title="Mes"),
            y=alt.Y("contrib_pct:Q", title="% contribución", axis=alt.Axis(format=".2f")),
            color=alt.condition(alt.datum.riesgo, alt.value("#ff6961"), alt.value("#1f77b4")),
            tooltip=[
                "mes",
                alt.Tooltip("monto", title="Devengado", format=",.2f"),
                alt.Tooltip("contrib_pct", title="Contrib. %", format=".2f"),
            ],
        )
        .properties(width=520, height=220)
    )
    st.altair_chart(chart, use_container_width=False)
    avance_series_display = avance_series.rename(columns={"contrib_pct": "contrib_%"})
    avance_series_display = round_numeric_for_reporting(avance_series_display)
    fmt_avance = build_style_formatters(avance_series_display)
    avance_style = avance_series_display.style.applymap(
        lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
        subset=["contrib_%"],
    )
    if fmt_avance:
        avance_style = avance_style.format(fmt_avance)
    st.dataframe(avance_style, use_container_width=True)

    if current_month < 12 and pim_total > 0:
        st.subheader("Proyección de ejecución por área (95% confianza)")
        if "sec_func" not in df_view.columns:
            st.info("Agrega la columna sec_func para proyectar la ejecución por área.")
        else:
            dev_sec = df_view.groupby("sec_func")[dev_cols].sum().reset_index()
            dev_sec_long = dev_sec.melt(id_vars="sec_func", var_name="col", value_name="monto")
            dev_sec_long["mes"] = dev_sec_long["col"].map({f"mto_devenga_{i:02d}": i for i in range(1, 13)})
            dev_sec_long = dev_sec_long.dropna(subset=["mes"])
            if dev_sec_long.empty:
                st.info("No hay datos históricos suficientes para proyectar la ejecución por área.")
            else:
                real_sec = dev_sec_long[dev_sec_long["mes"] <= current_month].copy()
                real_sec["sec_func"] = real_sec["sec_func"].astype(str)
                if real_sec.empty:
                    st.info("No hay registros devengados en los meses analizados para generar la proyección.")
                else:
                    remaining_months = max(12 - current_month, 0)
                    unique_areas = sorted(df_view["sec_func"].dropna().astype(str).unique().tolist())
                    stats = real_sec.groupby("sec_func")["monto"].agg(["count", "mean", "std"])
                    stats = stats.reindex(unique_areas).fillna(0.0)
                    stats.rename(columns={"count": "n", "mean": "mean", "std": "std"}, inplace=True)
                    stats["per_mes_proj"] = stats.apply(
                        lambda row: row["mean"] if row["n"] <= 1 or row["std"] == 0 else max(
                            row["mean"] - Z_SCORE_95 * row["std"] / np.sqrt(row["n"]),
                            0.0,
                        ),
                        axis=1,
                    )
                    proj_records = []
                    for area, row in stats.iterrows():
                        per_month = row["per_mes_proj"]
                        if remaining_months > 0:
                            for m in range(current_month + 1, 13):
                                proj_records.append(
                                    {"sec_func": area, "mes": m, "monto": per_month, "tipo": "Necesario"}
                                )
                    real_sec["tipo"] = "Real"
                    proj_sec = pd.DataFrame(proj_records)
                    dev_proj_sec = pd.concat(
                        [real_sec[["sec_func", "mes", "monto", "tipo"]], proj_sec],
                        ignore_index=True,
                    )
                    if not dev_proj_sec.empty:
                        dev_proj_sec["monto"] = dev_proj_sec["monto"].round(2)
                        chart_proj = (
                            alt.Chart(dev_proj_sec)
                            .mark_bar()
                            .encode(
                                x=alt.X("mes:O", title="Mes"),
                                y=alt.Y("monto:Q", title="Devengado"),
                                color=alt.Color("sec_func:N", title="Área"),
                                column=alt.Column("tipo:N", title=""),
                                tooltip=["sec_func", "mes", alt.Tooltip("monto", format=",.2f")],
                            )
                            .properties(width=160, height=240)
                        )
                        st.altair_chart(chart_proj, use_container_width=True)
                        proyeccion_wide = (
                            dev_proj_sec.pivot_table(
                                index="mes", columns=["sec_func", "tipo"], values="monto", fill_value=0
                            )
                            .sort_index(axis=1)
                            .reset_index()
                        )
                        proyeccion_wide.columns = ["mes"] + [f"{sec}_{tipo}" for sec, tipo in proyeccion_wide.columns[1:]]
                        proyeccion_wide = round_numeric_for_reporting(proyeccion_wide)

                        pim_sec = df_view.groupby("sec_func")["mto_pim"].sum().reindex(unique_areas, fill_value=0.0)
                        real_totals = real_sec.groupby("sec_func")["monto"].sum().reindex(unique_areas, fill_value=0.0)
                        proyectado_futuro = stats["per_mes_proj"] * remaining_months
                        total_estimado = real_totals + proyectado_futuro
                        avance_estimado = np.where(pim_sec > 0, total_estimado / pim_sec * 100.0, 0.0)

                        proyeccion_resumen = pd.DataFrame(
                            {
                                "sec_func": unique_areas,
                                "devengado_real": real_totals.values,
                                "proyeccion_futura": proyectado_futuro.values,
                                "total_estimado": total_estimado.values,
                                "mto_pim": pim_sec.values,
                                "avance_estimado_%": avance_estimado,
                            }
                        )
                        proyeccion_resumen = round_numeric_for_reporting(proyeccion_resumen)
                        fmt_proj = build_style_formatters(proyeccion_resumen)
                        resumen_style = proyeccion_resumen.style
                        if "avance_estimado_%" in proyeccion_resumen.columns:
                            resumen_style = resumen_style.applymap(
                                lambda v: "background-color: #ffcccc" if v < float(meta_avance) else "",
                                subset=["avance_estimado_%"],
                            )
                        if fmt_proj:
                            resumen_style = resumen_style.format(fmt_proj)
                        st.dataframe(resumen_style, use_container_width=True)

ritmo_df = pd.DataFrame()
leaderboard_df = pd.DataFrame()
if "mto_pim" in df_view.columns:
    st.subheader("Ritmo requerido por proceso")
    remaining_months = max(12 - current_month, 1)
    pim_total = df_view["mto_pim"].sum()
    processes = []
    for col, label in [("mto_certificado", "Certificar"), ("mto_compro_anual", "Comprometer"), ("devengado", "Devengar")]:
        total = df_view.get(col, pd.Series(dtype=float)).sum()
        actual_avg = total / current_month
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

# =========================
# Top sec_func con menor avance (leaderboard)
# =========================
if "sec_func" in df_view.columns and "mto_pim" in df_view.columns:
    st.subheader("Top áreas con menor avance")
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
        st.info("No hay datos disponibles para calcular el rendimiento por área.")

# =========================
# Gestión de alertas por correo
# =========================
st.subheader("Automatización de alertas por Outlook")
st.markdown(
    "Configura destinatarios, redacta una plantilla y envía correos de alerta para las áreas con avance por debajo del umbral."
)

if "alert_contacts" not in st.session_state:
    st.session_state["alert_contacts"] = {}

alert_df = leaderboard_df.copy()
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
        st.info(
            "Faltan correos para: " + ", ".join(missing_contacts)
        )

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

# =========================
# Descarga a Excel
# =========================
buf = to_excel_download(
    resumen=round_numeric_for_reporting(pivot.copy()),
    avance=round_numeric_for_reporting(avance_series.rename(columns={"contrib_pct": "contrib_%"})),
    proyeccion=proyeccion_wide,
    ritmo=round_numeric_for_reporting(ritmo_df.copy()),
    leaderboard=round_numeric_for_reporting(leaderboard_df.copy()),
)
st.download_button(
    "Descargar Excel (Resumen + Avance)",
    data=buf,
    file_name="siaf_resumen_avance.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
