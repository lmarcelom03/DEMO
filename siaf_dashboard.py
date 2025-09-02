import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="SIAF Dashboard - Peru Compras", layout="wide")
st.title("📊 SIAF Dashboard - Peru Compras")

st.markdown(
    "Carga el **Excel SIAF** y genera resúmenes de PIA, PIM, Certificado, Comprometido, **Devengado**, Girado y Pagado. "
    "Lee por defecto **A:EC** (base hasta CH + pasos CI-EC). "
    "Construye el **clasificador concatenado** normalizado para que siempre comience con `2.`."
)

# ---- Sidebar ----
with st.sidebar:
    st.header("⚙️ Parámetros de lectura")
    uploaded = st.file_uploader("Archivo SIAF (.xlsx)", type=["xlsx"])
    usecols = st.text_input("Rango de columnas (Excel)", "A:EC")
    sheet_name = st.text_input("Nombre de hoja (opcional)", "")
    header_row_excel = st.number_input("Fila de encabezados (Excel, 1=primera)", min_value=1, value=4)
    detect_header = st.checkbox("Autodetectar encabezado", value=True)
    st.markdown("---")
    st.header("🧮 Reglas CI-EC")
    current_month = st.number_input("Mes actual (1-12)", min_value=1, max_value=12, value=9)
    riesgo_umbral = st.number_input("Umbral de avance mínimo (%)", min_value=0, max_value=100, value=60)

# ---- Funciones auxiliares ----
def autodetect_sheet_and_header(xls, excel_bytes, usecols, user_sheet, header_guess):
    candidate_sheets = [user_sheet] if user_sheet else xls.sheet_names
    for s in candidate_sheets:
        try:
            tmp = pd.read_excel(excel_bytes, sheet_name=s, header=None, usecols=usecols, nrows=12)
        except Exception:
            continue
        for r in range(min(8, len(tmp))):
            row_vals = tmp.iloc[r].astype(str).str.lower().tolist()
            hits = sum(int(any(k in v for k in ["ano_eje","pim","pia","mto_","devenga","girado"])) for v in row_vals)
            if hits >= 2:
                return s, r
    return xls.sheet_names[0], header_guess-1

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
    return [f"{prefix}{i:02d}" for i in range(1,13) if f"{prefix}{i:02d}" in df.columns]

def ensure_ci_ec_steps(df, month, umbral):
    df = df.copy()
    dev_cols = find_monthly_columns(df, "mto_devenga_")
    if "devengado" not in df.columns:
        df["devengado"] = df[dev_cols].sum(axis=1) if dev_cols else 0.0
    col_mes = f"mto_devenga_{int(month):02d}"
    if "devengado_mes" not in df.columns:
        df["devengado_mes"] = df[col_mes] if col_mes in df.columns else 0.0
    if "saldo_pim" not in df.columns:
        df["saldo_pim"] = np.where(df.get("mto_pim",0)>0, df["mto_pim"]-df["devengado"],0.0)
    if "avance_%" not in df.columns:
        df["avance_%"] = np.where(df.get("mto_pim",0)>0, df["devengado"]/df["mto_pim"]*100.0,0.0)
    if "riesgo_devolucion" not in df.columns:
        df["riesgo_devolucion"] = df["avance_%"] < float(umbral)
    if "area" not in df.columns:
        df["area"] = ""
    return df

# ---- Clasificador ----
_code_re = re.compile(r"^\s*(\d+(?:\.\d+)*)")
def extract_code(text):
    if pd.isna(text): return ""
    s = str(text).strip()
    m = _code_re.match(s)
    return m.group(1) if m else ""

def last_segment(code):
    return code.split(".")[-1] if code else ""

def concat_hierarchy(gen, sub, subdet, esp, espdet):
    parts = []
    if gen: parts.append(gen)
    for child in [sub, subdet, esp, espdet]:
        if not child: continue
        if parts and (child.startswith(parts[-1] + ".") or child.startswith(parts[0] + ".")):
            parts.append(child)
        else:
            if parts:
                parts.append(parts[-1] + "." + last_segment(child))
            else:
                parts.append(child)
    return parts[-1] if parts else ""

def normalize_clasificador(code):
    if not code: return "2."
    return code if code.startswith("2.") else "2." + code

def build_classifier_columns(df):
    df = df.copy()
    gen = df.get("generica","")
    sub = df.get("subgenerica","")
    subdet = df.get("subgenerica_det","")
    esp = df.get("especifica","")
    espdet = df.get("especifica_det","")

    df["gen_cod"] = gen.map(extract_code) if "generica" in df.columns else ""
    df["sub_cod"] = sub.map(extract_code) if "subgenerica" in df.columns else ""
    df["subdet_cod"] = subdet.map(extract_code) if "subgenerica_det" in df.columns else ""
    df["esp_cod"] = esp.map(extract_code) if "especifica" in df.columns else ""
    df["espdet_cod"] = espdet.map(extract_code) if "especifica_det" in df.columns else ""

    # Aquí se arma el clasificador_cod y se normaliza con 2.
    df["clasificador_cod"] = [
        normalize_clasificador(concat_hierarchy(g,s,sd,e,ed))
        for g,s,sd,e,ed in zip(df["gen_cod"],df["sub_cod"],df["subdet_cod"],df["esp_cod"],df["espdet_cod"])
    ]
    return df

# ---- Pivot ----
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
    g = df.groupby(group_col,dropna=False)[cols].sum().reset_index()
    if "mto_pim" in g.columns and "devengado" in g.columns:
        g["saldo_pim"] = g["mto_pim"]-g["devengado"]
        g["avance_%"] = np.where(g["mto_pim"]>0,g["devengado"]/g["mto_pim"]*100.0,0.0)
    return g

# ---- Carga de archivo ----
if uploaded is None:
    st.info("Sube tu archivo Excel SIAF para empezar. Usa A:EC.")
    st.stop()

df, used_sheet = load_data(uploaded, usecols, sheet_name.strip() or None, int(header_row_excel), autodetect=detect_header)
st.success(f"Leída la hoja {used_sheet} con {df.shape[0]} filas y {df.shape[1]} columnas.")

# ---- Procesos ----
df_proc = ensure_ci_ec_steps(df, current_month, riesgo_umbral)
df_proc = build_classifier_columns(df_proc)

# ---- Vista rápida ----
st.subheader("Procesos CI-EC (con clasificador normalizado)")
st.dataframe(df_proc[["clasificador_cod","mto_pim","devengado","saldo_pim","avance_%"]].head(50))
