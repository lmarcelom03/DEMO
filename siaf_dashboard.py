
import re
import io
import json
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="SIAF Dashboard - Peru Compras", layout="wide")
st.title("📊 SIAF Dashboard - Peru Compras")

st.markdown(
    "Carga el **Excel SIAF** y genera resúmenes de PIA, PIM, Certificado, Comprometido, **Devengado**, Girado y Pagado. "
    "Ademas, si tu archivo tiene columnas **desde CI hasta EC**, el app puede **replicar y/o recalcular** esos pasos. "
    "Ahora tambien construye el **clasificador concatenado**: `generica.subgenerica.subgenerica_det.especifica.especifica_det`."
)

# ---- Sidebar: parametros ----
with st.sidebar:
    st.header("⚙️ Parametros de lectura")
    uploaded = st.file_uploader("Archivo SIAF (.xlsx)", type=["xlsx"])
    usecols = st.text_input("Rango de columnas (Excel)", "A:EC", help="Base A:CH + calculos CI:EC, como pediste.")
    sheet_name = st.text_input("Nombre de hoja (opcional)", "", help="Dejalo vacio para autodetectar la hoja de datos crudos.")
    header_row_excel = st.number_input(
        "Fila de encabezados (Excel, 1=primera)", min_value=1, value=4,
        help="Por defecto 4 (encabezados en fila 4)."
    )
    detect_header = st.checkbox("Autodetectar encabezado", value=True)
    st.markdown("---")
    st.header("🧮 Reglas CI-EC")
    st.caption("Si no existen en tu archivo, las calculamos aqui.")
    current_month = st.number_input("Mes actual (1-12)", min_value=1, max_value=12, value=9)
    riesgo_umbral = st.number_input("Umbral de avance minimo (%)", min_value=0, max_value=100, value=60)
    st.caption("Se marca riesgo_devolucion cuando Avance% < Umbral.")

def autodetect_sheet_and_header(xls, excel_bytes, usecols, user_sheet, header_guess):
    candidate_sheets = [user_sheet] if user_sheet else xls.sheet_names
    best = None
    for s in candidate_sheets:
        try:
            tmp = pd.read_excel(excel_bytes, sheet_name=s, header=None, usecols=usecols, nrows=12)
        except Exception:
            continue
        score_row = -1
        score_val = -1
        for r in range(min(8, len(tmp))):
            row_vals = tmp.iloc[r].astype(str).str.lower().tolist()
            hits = sum(int(any(k in v for k in ["ano_eje", "pim", "pia", "mto_", "devenga", "girado"])) for v in row_vals)
            if hits > score_val:
                score_val = hits
                score_row = r
        if score_val >= 2:
            best = (s, score_row)
            break
    if best is None:
        fallback = []
        for s in xls.sheet_names:
            try:
                df = pd.read_excel(excel_bytes, sheet_name=s, header=header_guess-1, usecols=usecols)
                fallback.append((s, df.shape[0], df.shape[1]))
            except Exception:
                continue
        if fallback:
            fallback.sort(key=lambda x: (x[2], x[1]), reverse=True)
            best = (fallback[0][0], header_guess-1)
        else:
            best = (xls.sheet_names[0], header_guess-1)
    return best

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

def find_monthly_columns(df, prefix):
    pats = [f"{prefix}{i:02d}" for i in range(1, 13)]
    return [c for c in pats if c in df.columns]

def ensure_ci_ec_steps(df, month, umbral):
    """Asegura/calcula campos equivalentes a las logicas CI-EC (genericas)."""
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

import pandas as pd

_code_re = re.compile(r"^\s*(\d+(?:\.\d+)*)")
def extract_code(text):
    if pd.isna(text): 
        return ""
    s = str(text).strip()
    m = _code_re.match(s)
    return m.group(1) if m else ""

def last_segment(code):
    if not code: return ""
    return code.split(".")[-1]

def concat_hierarchy(gen, sub, subdet, esp, espdet):
    parts = []
    if gen: parts.append(gen)
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
    full = parts[-1] if parts else ""
    return full

def build_classifier_columns(df):
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
    def normalize_clasificador(code):
    if not code:
        return "2."
    return code if code.startswith("2.") else "2." + code

    df["clasificador_cod"] = [
        concat_hierarchy(g, s, sd, e, ed)
        for g, s, sd, e, ed in zip(df["gen_cod"], df["sub_cod"], df["subdet_cod"], df["esp_cod"], df["espdet_cod"])
    ]

    def desc(text):
        if pd.isna(text): return ""
        s = str(text)
        return s.split(".", 1)[1].strip() if "." in s else s

    df["generica_desc"] = gen.map(desc) if "generica" in df.columns else ""
    df["subgenerica_desc"] = sub.map(desc) if "subgenerica" in df.columns else ""
    df["subgenerica_det_desc"] = subdet.map(desc) if "subgenerica_det" in df.columns else ""
    df["especifica_desc"] = esp.map(desc) if "especifica" in df.columns else ""
    df["especifica_det_desc"] = espdet.map(desc) if "especifica_det" in df.columns else ""

    df["clasificador_desc"] = (
        df["generica_desc"].fillna("")
        + " > " + df["subgenerica_desc"].fillna("")
        + " > " + df["subgenerica_det_desc"].fillna("")
        + " > " + df["especifica_desc"].fillna("")
        + " > " + df["especifica_det_desc"].fillna("")
    ).str.strip(" >")

    return df

def pivot_exec(df, group_col, dev_cols):
    cols = []
    if "mto_pia" in df.columns: cols.append("mto_pia")
    if "mto_pim" in df.columns: cols.append("mto_pim")
    if "mto_certificado" in df.columns: cols.append("mto_certificado")
    if "mto_compro_anual" in df.columns: cols.append("mto_compro_anual")
    if dev_cols: cols.append("devengado")
    if "devengado" not in df.columns and dev_cols:
        df = df.copy()
        df["devengado"] = df[dev_cols].sum(axis=1)
    g = df.groupby(group_col, dropna=False)[cols].sum().reset_index()
    if "mto_pim" in g.columns and "devengado" in g.columns:
        g["saldo_pim"] = g["mto_pim"] - g["devengado"]
        g["avance_%"] = np.where(g["mto_pim"] > 0, g["devengado"] / g["mto_pim"] * 100.0, 0.0)
    return g

def to_excel_download(**dfs):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, d in dfs.items():
            d.to_excel(writer, index=False, sheet_name=name[:31] or "Sheet1")
    output.seek(0)
    return output

# ---- Carga ----
if uploaded is None:
    st.info("Sube tu archivo Excel SIAF para empezar. Usa A:EC si quieres incluir los pasos CI-EC.")
    st.stop()

excel_bytes = uploaded
try:
    df, used_sheet = load_data(excel_bytes, usecols, sheet_name.strip() or None, int(header_row_excel), autodetect=detect_header)
except Exception as e:
    st.error(f"No pude leer el archivo: {e}")
    st.stop()

st.success(f"Leida la hoja **{used_sheet}** con **{df.shape[0]}** filas y **{df.shape[1]}** columnas.")

# ---- Filtros
st.subheader("🔍 Filtros")
filter_cols = [c for c in df.columns if any(k in c for k in [
    "unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func","departamento_meta","provincia_meta","distrito_meta"
])]

cols1, cols2, cols3 = st.columns(3)
selected_filters = {}
for i, c in enumerate(filter_cols):
    with (cols1 if i % 3 == 0 else cols2 if i % 3 == 1 else cols3):
        vals = sorted([str(x) for x in df[c].dropna().unique().tolist()])
        if len(vals) > 1:
            chosen = st.multiselect(c, options=vals, default=[])
            if chosen:
                selected_filters[c] = set(chosen)

mask = pd.Series([True] * len(df))
for c, allowed in selected_filters.items():
    mask &= df[c].astype(str).isin(allowed)
df_f = df[mask].copy()

# ---- CI-EC y clasificador ----
df_ci_ec = ensure_ci_ec_steps(df_f, current_month, riesgo_umbral)
df_ci_ec = build_classifier_columns(df_ci_ec)

# ---- Resumen ejecutivo
st.subheader("📌 Resumen ejecutivo (totales)")
dev_cols = [c for c in df_ci_ec.columns if c.startswith("mto_devenga_")]
tot_pia = float(df_ci_ec.get("mto_pia", 0).sum())
tot_pim = float(df_ci_ec.get("mto_pim", 0).sum())
tot_dev = float(df_ci_ec.get("devengado", 0).sum())
tot_cert = float(df_ci_ec.get("mto_certificado", 0).sum()) if "mto_certificado" in df_ci_ec.columns else 0.0
tot_comp = float(df_ci_ec.get("mto_compro_anual", 0).sum()) if "mto_compro_anual" in df_ci_ec.columns else 0.0
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

# ---- Agrupaciones
st.subheader("📈 Vistas por agrupacion")
group_options = [c for c in df_ci_ec.columns if c in [
    "clasificador_cod","unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func","area"
]]
default_idx = group_options.index("clasificador_cod") if "clasificador_cod" in group_options else 0
group_col = st.selectbox("Agrupar por", options=group_options, index=default_idx)
pivot = pivot_exec(df_ci_ec, group_col, dev_cols)
st.dataframe(pivot, use_container_width=True)

# ---- Procesos CI-EC: detalle por clasificador
st.subheader("🧩 Procesos CI-EC (monto vinculado a su categoria)")
ci_cols_show = [
    "clasificador_cod", "generica", "subgenerica", "subgenerica_det", "especifica", "especifica_det",
    "mto_pia","mto_pim","mto_certificado","mto_compro_anual","devengado_mes","devengado","saldo_pim","avance_%","riesgo_devolucion"
]
ci_cols_show = [c for c in ci_cols_show if c in df_ci_ec.columns]
st.caption("Cada monto queda enlazado a su cadena: generica.subgenerica.subgenerica_det.especifica.especifica_det")
st.dataframe(df_ci_ec[ci_cols_show].head(50), use_container_width=True)

# ---- Tabla consolidada por clasificador (suma)
agg_cols = [c for c in ["mto_pia","mto_pim","mto_certificado","mto_compro_anual","devengado_mes","devengado","saldo_pim"] if c in df_ci_ec.columns]
consolidado = df_ci_ec.groupby(["clasificador_cod","generica","subgenerica","subgenerica_det","especifica","especifica_det"], dropna=False)[agg_cols].sum().reset_index()
if "mto_pim" in consolidado.columns and "devengado" in consolidado.columns:
    consolidado["avance_%"] = np.where(consolidado["mto_pim"] > 0, consolidado["devengado"]/consolidado["mto_pim"]*100.0, 0.0)
st.markdown("**Consolidado por clasificador**")
st.dataframe(consolidado.head(200), use_container_width=True)

# ---- Serie mensual
if dev_cols:
    st.subheader("🗓️ Devengado mensual (por filtro)")
    month_map = {f"mto_devenga_{i:02d}": i for i in range(1,13)}
    dev_series = df_ci_ec[dev_cols].sum().reset_index()
    dev_series.columns = ["col", "monto"]
    dev_series["mes"] = dev_series["col"].map(month_map)
    dev_series = dev_series.sort_values("mes")
    st.dataframe(dev_series[["mes","monto"]])

    fig, ax = plt.subplots()
    ax.bar(dev_series["mes"].astype(int), dev_series["monto"].values)
    ax.set_xlabel("Mes")
    ax.set_ylabel("Devengado (S/)")
    ax.set_title("Devengado mensual (acumulado por filtro)")
    st.pyplot(fig)

# ---- Descarga
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df_ci_ec.to_excel(writer, index=False, sheet_name="Datos_CI_EC")
    pivot.to_excel(writer, index=False, sheet_name="Resumen")
    consolidado.to_excel(writer, index=False, sheet_name="Consolidado_Clasificador")
buf.seek(0)
st.download_button("⬇️ Descargar Excel (CI-EC + Resumen + Clasificador)", data=buf, file_name="siaf_ci_ec_clasificador.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
