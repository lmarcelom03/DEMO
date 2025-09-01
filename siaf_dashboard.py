
import re
import io
import json
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="SIAF Dashboard - Perú Compras", layout="wide")

st.title("📊 SIAF Dashboard – Perú Compras")

st.markdown(
    "Sube el **Excel del SIAF** (A:CH) y obtén resúmenes de **PIA, PIM, Certificado, Comprometido, Devengado, Girado y Pagado** "
    "con filtros por **UE, programa, función, fuente, genérica, específica, meta**, etc. "
)

# ---- Sidebar: carga y parámetros ----
with st.sidebar:
    st.header("⚙️ Parámetros de lectura")
    uploaded = st.file_uploader("Archivo SIAF (Excel .xlsx)", type=["xlsx"])
    usecols = st.text_input("Rango de columnas (Excel)", "A:CH", help="Por defecto A:CH como mencionaste.")
    sheet_name = st.text_input("Nombre de hoja (opcional)", "", help="Si se deja vacío, se autodetecta la hoja con 'ano_eje' y muchas columnas.")
    header_row_excel = st.number_input(
        "Fila de encabezados (Excel, 1=primera)", min_value=1, value=4,
        help="Tu caso típico es encabezados en la fila 4. Se intentará autodetección si no coincide."
    )
    detect_header = st.checkbox("Autodetectar encabezado", value=True, help="Usa heurística para encontrar la fila que contiene 'ano_eje', 'mto_pim', etc.")
    st.markdown("---")
    st.caption("Consejo: si tu archivo tiene una hoja llamada **Hoja1** con los datos crudos, déjala seleccionada o en autodetección.")

def autodetect_sheet_and_header(xls, excel_bytes, usecols, user_sheet, header_guess):
    # 1) Try user sheet first (if any)
    candidate_sheets = [user_sheet] if user_sheet else xls.sheet_names
    best = None
    for s in candidate_sheets:
        try:
            # Read a few rows without header to scan
            tmp = pd.read_excel(excel_bytes, sheet_name=s, header=None, usecols=usecols, nrows=12)
        except Exception:
            continue
        # find a row that looks like header: contains 'ano_eje' or several known tokens
        score_row = -1
        score_val = -1
        for r in range(min(8, len(tmp))):
            row_vals = tmp.iloc[r].astype(str).str.lower().tolist()
            hits = sum(int(any(k in v for k in ["ano_eje", "pim", "pia", "mto_", "devenga", "girado"])) for v in row_vals)
            if hits > score_val:
                score_val = hits
                score_row = r
        if score_val >= 2:  # looks good
            best = (s, score_row)
            break
    if best is None:
        # fallback: try all sheets to pick the one with most columns when header at user guess-1
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

    # Clean
    df = df.dropna(how="all").dropna(axis=1, how="all")
    # normalize colnames
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df, s

def find_monthly_columns(df, prefix):
    # e.g., prefix = "mto_devenga_"
    pats = [f"{prefix}{i:02d}" for i in range(1, 13)]
    return [c for c in pats if c in df.columns]

def smart_numeric(df, cols):
    return df[cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

def build_summary(df):
    # metric columns
    base_cols = {
        "pia": "mto_pia",
        "pim": "mto_pim",
        "certificado": "mto_certificado",
        "comprometido": "mto_compro_anual",
    }
    present = {k: v for k, v in base_cols.items() if v in df.columns}

    # monthly devengado / girado / pagado
    dev_cols = find_monthly_columns(df, "mto_devenga_")
    gir_cols = find_monthly_columns(df, "mto_girado_")
    pag_cols = find_monthly_columns(df, "mto_pagado_")

    # numeric
    num = {}
    for k, col in present.items():
        num[k] = df[col].sum()

    if dev_cols:
        num["devengado"] = df[dev_cols].sum().sum()
    else:
        num["devengado"] = 0.0
    if gir_cols:
        num["girado"] = df[gir_cols].sum().sum()
    else:
        num["girado"] = 0.0
    if pag_cols:
        num["pagado"] = df[pag_cols].sum().sum()
    else:
        num["pagado"] = 0.0

    pim = num.get("pim", 0.0)
    dev = num.get("devengado", 0.0)
    saldo = pim - dev if pim else 0.0
    avance = (dev / pim * 100.0) if pim else 0.0

    num["saldo_pim"] = saldo
    num["avance_%"] = avance
    return num, dev_cols, gir_cols, pag_cols

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

if uploaded is None:
    st.info("Sube tu archivo Excel SIAF para empezar. Puedes usar **A:CH** y **fila 4** como mencionaste.")
    st.stop()

# Cargar datos
excel_bytes = uploaded
try:
    df, used_sheet = load_data(excel_bytes, usecols, sheet_name.strip() or None, int(header_row_excel), autodetect=detect_header)
except Exception as e:
    st.error(f"No pude leer el archivo: {e}")
    st.stop()

st.success(f"Leída la hoja **{used_sheet}** con **{df.shape[0]}** filas y **{df.shape[1]}** columnas.")

# ---- Filtros dinámicos ----
st.subheader("🔍 Filtros")
filter_cols = [c for c in df.columns if any(k in c for k in [
    "unidad_ejecutora", "fuente_financ", "generica", "subgenerica", "subgenerica_det",
    "especifica", "especifica_det", "funcion", "division_fn", "grupo_fn",
    "programa_pptal", "producto_proyecto", "activ_obra_accinv", "meta", "sec_func",
    "departamento_meta", "provincia_meta", "distrito_meta"
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

# aplicar filtros
mask = pd.Series([True] * len(df))
for c, allowed in selected_filters.items():
    mask &= df[c].astype(str).isin(allowed)
df_f = df[mask].copy()

# ---- Resumen ejecutivo ----
st.subheader("📌 Resumen ejecutivo (totales)")
summary, dev_cols, gir_cols, pag_cols = build_summary(df_f)

kpi_cols = st.columns(6)
kpis = [
    ("PIA", "pia"),
    ("PIM", "pim"),
    ("Certificado", "certificado"),
    ("Comprometido", "comprometido"),
    ("Devengado (YTD)", "devengado"),
    ("Saldo PIM", "saldo_pim"),
]
for i, (label, key) in enumerate(kpis):
    with kpi_cols[i % 6]:
        val = summary.get(key, 0.0)
        st.metric(label, f"S/ {val:,.2f}")

st.metric("Avance de ejecución", f"{summary.get('avance_%', 0.0):.2f}%")

# ---- Vistas por agrupación ----
st.subheader("📈 Vistas por agrupación")
group_options = [c for c in df_f.columns if c in [
    "unidad_ejecutora","fuente_financ","generica","subgenerica","subgenerica_det",
    "especifica","especifica_det","funcion","division_fn","grupo_fn","programa_pptal",
    "producto_proyecto","activ_obra_accinv","meta","sec_func"
]]
group_col = st.selectbox("Agrupar por", options=group_options, index=group_options.index("generica") if "generica" in group_options else 0)
pivot = pivot_exec(df_f, group_col, dev_cols)

st.dataframe(pivot)

# descarga de pivote y datos filtrados
dl1, dl2 = st.columns(2)
with dl1:
    buf = to_excel_download(Resumen=pivot, Datos_filtrados=df_f)
    st.download_button("⬇️ Descargar Excel (resumen + datos filtrados)", data=buf, file_name="siaf_resumen.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---- Serie mensual de devengado ----
if dev_cols:
    st.subheader("🗓️ Serie mensual – Devengado")
    # armar serie
    month_map = {f"mto_devenga_{i:02d}": i for i in range(1,13)}
    dev_series = df_f[dev_cols].sum().reset_index()
    dev_series.columns = ["col", "monto"]
    dev_series["mes"] = dev_series["col"].map(month_map)
    dev_series = dev_series.sort_values("mes")

    # Mostrar como tabla y gráfico
    st.dataframe(dev_series[["mes","monto"]])

    # gráfico simple con matplotlib (para evitar dependencias)
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()
    ax.bar(dev_series["mes"].astype(int), dev_series["monto"].values)
    ax.set_xlabel("Mes")
    ax.set_ylabel("Devengado (S/)")
    ax.set_title("Devengado mensual (acumulado por filtro)")
    st.pyplot(fig)

st.caption("Tip: usa los filtros de la parte superior para ver el **avance %** por UE, genérica, específica, fuente, etc. "
           "El botón de descarga incluye el resumen y los datos filtrados.")
