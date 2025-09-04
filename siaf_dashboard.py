import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="SIAF Dashboard - Peru Compras", layout="wide")
st.title("SIAF Dashboard - Peru Compras")

st.markdown(
    "Carga el **Excel SIAF** y obtén resúmenes de **PIA, PIM, Certificado, Comprometido, Devengado, Saldo PIM y % Avance**. "
    "Lee por defecto **A:EC** (base hasta CH + pasos CI–EC). "
    "Construye el **clasificador concatenado** y lo **normaliza para que siempre comience con `2.`**; además agrega una **descripción jerárquica**. "
    "Incluye filtros, pivotes, serie mensual y descarga a Excel."
)

# =========================
# Sidebar / parámetros
# =========================
with st.sidebar:
    st.header("Parámetros de lectura")
    uploaded = st.file_uploader("Archivo SIAF (.xlsx)", type=["xlsx"])
    usecols = st.text_input("Rango de columnas (Excel)", "A:EC", help="Recomendado para incluir CI–EC.")
    sheet_name = st.text_input("Nombre de hoja (opcional)", "", help="Déjalo vacío para autodetección.")
    header_row_excel = st.number_input("Fila de encabezados (Excel, 1=primera)", min_value=1, value=4)
    detect_header = st.checkbox("Autodetectar encabezado", value=True)
    st.markdown("---")
    st.header("Reglas CI–EC")
    current_month = st.number_input("Mes actual (1-12)", min_value=1, max_value=12, value=9)
    riesgo_umbral = st.number_input("Umbral de avance mínimo (%)", min_value=0, max_value=100, value=60)
    meta_avance = st.number_input("Meta de avance al cierre (%)", min_value=0, max_value=100, value=95)
    st.caption("Se marca riesgo_devolucion si Avance% < Umbral.")

# =========================
# Utilitarios de carga
# =========================
def autodetect_sheet_and_header(xls, excel_bytes, usecols, user_sheet, header_guess):
    """
    Busca la hoja y la fila que luce como encabezado (contenga 'ano_eje', 'mto_', 'pim', etc.).
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

def to_excel_download(resumen, avance, proyeccion=None, ritmo=None):
    wb = Workbook()
    ws_res = wb.active
    ws_res.title = "Resumen"
    for r in dataframe_to_rows(resumen, index=False, header=True):
        ws_res.append(r)
    for cell in ws_res[1]:
        cell.font = Font(bold=True)

    ws_av = wb.create_sheet("Avance")
    for r in dataframe_to_rows(avance, index=False, header=True):
        ws_av.append(r)

    chart = BarChart()
    data = Reference(ws_av, min_col=3, min_row=1, max_row=ws_av.max_row, max_col=3)
    cats = Reference(ws_av, min_col=1, min_row=2, max_row=ws_av.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.title = "Contribución mensual (%)"
    chart.y_axis.title = "%"
    chart.x_axis.title = "Mes"
    chart.height = 7
    chart.width = 15
    ws_av.add_chart(chart, "E2")

    if proyeccion is not None and not proyeccion.empty:
        ws_proj = wb.create_sheet("Proyeccion")
        for r in dataframe_to_rows(proyeccion, index=False, header=True):
            ws_proj.append(r)
        chart2 = BarChart()
        data2 = Reference(ws_proj, min_col=2, min_row=1, max_row=ws_proj.max_row, max_col=ws_proj.max_column)
        cats2 = Reference(ws_proj, min_col=1, min_row=2, max_row=ws_proj.max_row)
        chart2.add_data(data2, titles_from_data=True)
        chart2.set_categories(cats2)
        chart2.title = "Proyeccion devengado"
        chart2.y_axis.title = "Monto"
        chart2.x_axis.title = "Mes"
        chart2.height = 7
        chart2.width = 15
        ws_proj.add_chart(chart2, "E2")

    if ritmo is not None and not ritmo.empty:
        ws_rit = wb.create_sheet("Ritmo")
        for r in dataframe_to_rows(ritmo, index=False, header=True):
            ws_rit.append(r)
        chart3 = BarChart()
        data3 = Reference(ws_rit, min_col=2, min_row=1, max_row=ws_rit.max_row, max_col=ws_rit.max_column)
        cats3 = Reference(ws_rit, min_col=1, min_row=2, max_row=ws_rit.max_row)
        chart3.add_data(data3, titles_from_data=True)
        chart3.set_categories(cats3)
        chart3.title = "Ritmo actual vs necesario"
        chart3.y_axis.title = "Monto"
        chart3.x_axis.title = "Proceso"
        chart3.height = 7
        chart3.width = 15
        ws_rit.add_chart(chart3, "E2")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# =========================
# Carga del archivo
# =========================
if uploaded is None:
    st.info("Sube tu archivo Excel SIAF para empezar. Usa A:EC para incluir pasos CI–EC.")
    st.stop()

try:
    df, used_sheet = load_data(uploaded, usecols, sheet_name.strip() or None, int(header_row_excel), autodetect=detect_header)
except Exception as e:
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

st.success(f"Leída la hoja '{used_sheet}' con {df.shape[0]} filas y {df.shape[1]} columnas.")

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

pivot = pivot_exec(df_proc, group_col, dev_cols)
pivot_display = pivot
if "avance_%" in pivot_display.columns:
    pivot_display = pivot_display.style.applymap(
        lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
        subset=["avance_%"],
    )
st.dataframe(pivot_display, use_container_width=True)

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
ci_cols = [c for c in ci_cols if c in df_proc.columns]
df_ci = df_proc[ci_cols].head(300)
if "avance_%" in df_ci.columns:
    df_ci = df_ci.style.applymap(
        lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
        subset=["avance_%"],
    )
st.dataframe(df_ci, use_container_width=True)

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
    if c in df_proc.columns
]
consolidado = df_proc.groupby(
    ["clasificador_cod","clasificador_desc","generica","subgenerica","subgenerica_det","especifica","especifica_det"],
    dropna=False
)[agg_cols].sum().reset_index()

if "mto_pim" in consolidado.columns and "devengado" in consolidado.columns:
    consolidado["avance_%"] = np.where(consolidado["mto_pim"] > 0, consolidado["devengado"]/consolidado["mto_pim"]*100.0, 0.0)

st.markdown("**Consolidado por clasificador**")
consol_display = consolidado.head(500)
if "avance_%" in consol_display.columns:
    consol_display = consol_display.style.applymap(
        lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
        subset=["avance_%"],
    )
st.dataframe(consol_display, use_container_width=True)

# =========================
# Serie mensual interactiva
# =========================
avance_series = pd.DataFrame()
proyeccion_wide = pd.DataFrame()
if dev_cols and "mto_pim" in df_proc.columns:
    st.subheader("Avance mensual interactivo")
    month_map = {f"mto_devenga_{i:02d}": i for i in range(1, 13)}
    dev_series = df_proc[dev_cols].sum().reset_index()
    dev_series.columns = ["col", "monto"]
    dev_series["mes"] = dev_series["col"].map(month_map)
    dev_series = dev_series.sort_values("mes")
    pim_total = df_proc["mto_pim"].sum()
    dev_series["contrib_pct"] = np.where(pim_total > 0, dev_series["monto"] / pim_total * 100.0, 0.0)
    dev_series["riesgo"] = dev_series["contrib_pct"] < float(riesgo_umbral)
    avance_series = dev_series[["mes", "monto", "contrib_pct"]]
    chart = (
        alt.Chart(dev_series)
        .mark_bar()
        .encode(
            x=alt.X("mes:O", title="Mes"),
            y=alt.Y("contrib_pct:Q", title="% contribución"),
            color=alt.condition(alt.datum.riesgo, alt.value("#ff6961"), alt.value("#1f77b4")),
            tooltip=[
                "mes",
                alt.Tooltip("monto", title="Devengado", format=","),
                alt.Tooltip("contrib_pct", title="Contrib. %", format=".2f"),
            ],
        )
        .properties(width=600, height=250)
    )
    st.altair_chart(chart, use_container_width=False)
    st.dataframe(
        avance_series.style.applymap(
            lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
            subset=["contrib_pct"],
        ),
        use_container_width=True,
    )

    if current_month < 12 and pim_total > 0:
        st.subheader("Proyección de ejecución")
        target_total = pim_total * float(meta_avance) / 100.0
        dev_acum = dev_series.loc[dev_series["mes"] <= current_month, "monto"].sum()
        remaining_needed = max(target_total - dev_acum, 0)
        remaining_months = 12 - current_month
        proj_months = list(range(current_month + 1, 13))
        proj_per_month = remaining_needed / remaining_months if remaining_months > 0 else 0
        proj_df = pd.DataFrame({"mes": proj_months, "monto": [proj_per_month] * len(proj_months), "tipo": "Necesario"})
        real_df = dev_series[["mes", "monto"]].copy()
        real_df["tipo"] = "Real"
        dev_proj = pd.concat([real_df, proj_df], ignore_index=True)
        chart_proj = (
            alt.Chart(dev_proj)
            .mark_bar()
            .encode(
                x=alt.X("mes:O", title="Mes"),
                y=alt.Y("monto:Q", title="Devengado"),
                color="tipo:N",
                tooltip=["mes", alt.Tooltip("monto", format=",")],
            )
            .properties(width=650, height=250)
        )
        st.altair_chart(chart_proj, use_container_width=False)
        proyeccion_wide = dev_proj.pivot_table(index="mes", columns="tipo", values="monto", fill_value=0).reset_index()

ritmo_df = pd.DataFrame()
if "mto_pim" in df_proc.columns:
    st.subheader("Ritmo requerido por proceso")
    remaining_months = max(12 - current_month, 1)
    pim_total = df_proc["mto_pim"].sum()
    processes = []
    for col, label in [("mto_certificado", "Certificar"), ("mto_compro_anual", "Comprometer"), ("devengado", "Devengar")]:
        total = df_proc.get(col, pd.Series(dtype=float)).sum()
        actual_avg = total / current_month
        needed = max(pim_total - total, 0)
        required_avg = needed / remaining_months
        processes.append({"Proceso": label, "Actual": actual_avg, "Necesario": required_avg})
    ritmo_df = pd.DataFrame(processes)
    ritmo_melt = ritmo_df.melt("Proceso", var_name="Tipo", value_name="Monto")
    chart_ritmo = (
        alt.Chart(ritmo_melt)
        .mark_bar()
        .encode(
            x=alt.X("Proceso:N"),
            y=alt.Y("Monto:Q"),
            color="Tipo:N",
            tooltip=["Proceso", "Tipo", alt.Tooltip("Monto", format=",")],
        )
        .properties(width=650, height=300)
    )
    st.altair_chart(chart_ritmo, use_container_width=False)

# =========================
# Descarga a Excel
# =========================
buf = to_excel_download(resumen=pivot, avance=avance_series, proyeccion=proyeccion_wide, ritmo=ritmo_df)
st.download_button(
    "Descargar Excel (Resumen + Avance)",
    data=buf,
    file_name="siaf_resumen_avance.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
