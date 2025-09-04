!pip install scikit-learn

import re
import io
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from sklearn.linear_model import LinearRegression # Using sklearn for a simple projection example

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
    st.caption("Se marca riesgo_devolucion si Avance% < Umbral.")
    st.markdown("---")
    st.header("Proyección")
    projection_method = st.selectbox("Método de proyección", options=["Promedio Mensual", "Regresión Lineal"])


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
    cert_cols = find_monthly_columns(df, "mto_certificado_")
    comp_cols = find_monthly_columns(df, "mto_compro_")

    if "devengado" not in df.columns:
        df["devengado"] = df[dev_cols].sum(axis=1) if dev_cols else 0.0
    if "certificado" not in df.columns:
        df["certificado"] = df[cert_cols].sum(axis=1) if cert_cols else 0.0
    if "comprometido" not in df.columns:
         df["comprometido"] = df[comp_cols].sum(axis=1) if comp_cols else 0.0


    col_mes_dev = f"mto_devenga_{int(month):02d}"
    if "devengado_mes" not in df.columns:
        df["devengado_mes"] = df[col_mes_dev] if col_mes_dev in df.columns else 0.0

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

# =========================
# Styling Function
# =========================
def highlight_risk(row, umbral):
    """Highlights rows where 'avance_%' is below the specified umbral."""
    if 'avance_%' in row and isinstance(row['avance_%'], (int, float)) and row['avance_%'] < umbral:
        return ['background-color: yellow'] * len(row)
    return [''] * len(row)

# =========================
# Chart Generation Functions
# =========================

def generate_monthly_series_chart(df, cols, title, ylabel, color):
    """Generates a monthly time series chart for specified columns."""
    # Determine the appropriate month map based on the prefix of the columns
    prefix = cols[0].split('_')[1] if cols else None
    if prefix:
        month_map = {f"mto_{prefix}_{i:02d}": i for i in range(1, 13) if f"mto_{prefix}_{i:02d}" in cols}
    else:
        month_map = {} # Empty map if no columns

    series_data = df[cols].sum().reset_index()
    series_data.columns = ["col", "monto"]
    series_data["mes"] = series_data["col"].map(month_map)
    series_data = series_data.sort_values("mes")

    fig, ax = plt.subplots(figsize=(12, 7))
    ax.bar(series_data["mes"].astype(int), series_data["monto"].values, color=color)
    ax.set_xlabel("Mes")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.set_xticks(range(1, 13)) # Ensure all months are shown on x-axis
    st.pyplot(fig)
    plt.close(fig) # Close the figure

    return fig, series_data # Return figure and data for download


def generate_projection_chart(dev_series, current_month, projection_method, total_pim):
    """Generates a chart with historical execution and future projection."""
    fig, ax = plt.subplots(figsize=(12, 7))

    # Historical data (cumulative)
    historical_dev_cumulative = dev_series[dev_series["mes"] <= current_month].copy()
    historical_dev_cumulative["cumulative_monto"] = historical_dev_cumulative["monto"].cumsum()

    ax.plot(historical_dev_cumulative["mes"].astype(int), historical_dev_cumulative["cumulative_monto"].values, color='skyblue', marker='o', label='Devengado Histórico Acumulado')
    ax.fill_between(historical_dev_cumulative["mes"].astype(int), historical_dev_cumulative["cumulative_monto"].values, color='skyblue', alpha=0.3)


    # Projection
    remaining_months = list(range(current_month + 1, 13))
    if remaining_months:
        last_historical_cumulative = historical_dev_cumulative["cumulative_monto"].iloc[-1] if not historical_dev_cumulative.empty else 0

        if projection_method == "Promedio Mensual":
            average_monthly_dev = historical_dev_cumulative["monto"].mean() if not historical_dev_cumulative.empty else 0
            projected_cumulative_values = [last_historical_cumulative + average_monthly_dev * i for i in range(1, len(remaining_months) + 1)]

        elif projection_method == "Regresión Lineal":
             if len(historical_dev_cumulative) >= 2:
                model = LinearRegression()
                X = historical_dev_cumulative["mes"].values.reshape(-1, 1)
                y = historical_dev_cumulative["cumulative_monto"].values
                model.fit(X, y)

                projection_months_array = np.array(remaining_months).reshape(-1, 1)
                projected_cumulative_values = model.predict(projection_months_array)

                # Ensure projected cumulative values are non-decreasing and above last historical value
                projected_cumulative_values = np.maximum.accumulate(projected_cumulative_values)
                projected_cumulative_values[projected_cumulative_values < last_historical_cumulative] = last_historical_cumulative


             else:
                # Fallback to average if not enough data points for regression
                average_monthly_dev = historical_dev_cumulative["monto"].mean() if not historical_dev_cumulative.empty else 0
                projected_cumulative_values = [last_historical_cumulative + average_monthly_dev * i for i in range(1, len(remaining_months) + 1)]


        if projected_cumulative_values:
            projection_months = remaining_months
            ax.plot(projection_months, projected_cumulative_values, color='orange', linestyle='--', marker='o', label=f'Proyección ({projection_method})')
            ax.fill_between(projection_months, projected_cumulative_values, last_historical_cumulative, color='orange', alpha=0.2)

    # Add total PIM line
    if total_pim > 0:
        ax.axhline(y=total_pim, color='red', linestyle='-', label=f'PIM Total (S/ {total_pim:,.2f})')


    ax.set_xlabel("Mes")
    ax.set_ylabel("Monto Acumulado (S/)")
    ax.set_title(f"Proyección de Ejecución Acumulada vs PIM")
    ax.set_xticks(range(1, 13))
    ax.legend()
    ax.grid(axis='y', linestyle='--')
    st.pyplot(fig)
    plt.close(fig)

    return fig # Return figure for download


# =========================
# Descarga a Excel (Modificada)
# =========================
def to_excel_download_modified(df_proc, pivot, consolidado, monthly_charts_figs, projection_fig, group_col, riesgo_umbral):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        workbook = writer.book

        # Add monthly charts images to sheets
        for chart_name, fig in monthly_charts_figs.items():
            if chart_name not in workbook.sheetnames:
                 workbook.create_sheet(chart_name)
            ws_chart = workbook[chart_name]

            img_buffer = io.BytesIO()
            fig.savefig(img_buffer, format='png')
            img_buffer.seek(0)
            # No need to close fig here, it's done in the generation function

            img = Image(img_buffer)
            ws_chart.add_image(img, 'A1')

        # Add projection chart image
        if projection_fig:
            if "Proyeccion Ejecucion" not in workbook.sheetnames:
                 workbook.create_sheet("Proyeccion Ejecucion")
            ws_proj = workbook["Proyeccion Ejecucion"]
            img_buffer = io.BytesIO()
            projection_fig.savefig(img_buffer, format='png')
            img_buffer.seek(0)
            # No need to close fig here, it's done in the generation function

            img = Image(img_buffer)
            ws_proj.add_image(img, 'A1')


        # Create and format the executive summary table based on pivot
        executive_summary_cols = [col for col in ['clasificador_cod', 'clasificador_desc', group_col, 'mto_pia', 'mto_pim', 'mto_certificado', 'mto_compro_anual', 'devengado', 'saldo_pim', 'avance_%'] if col in pivot.columns]
        executive_summary_df = pivot[executive_summary_cols].copy()

        # Apply highlighting to the executive summary table data before writing
        # Note: Styler formatting is not directly exported to Excel.
        # We export the raw data.
        executive_summary_df.to_excel(writer, index=False, sheet_name="Resumen Ejecutivo")


    output.seek(0)
    return output

# =========================
# Main App Logic
# =========================

if uploaded is not None:
    # Data loading and initial processing are done above

    # ... (Filtering and CI-EC/Clasificador processing also done above)

    # Display Resumen Ejecutivo, Vistas por agrupación, Procesos CI–EC, Consolidado
    # (Code for these sections is already in the initial code block and should be kept)

    # =========================
    # Monthly Series Charts
    # =========================
    st.subheader("Ritmo Mensual: Certificado, Comprometido y Devengado")

    # Ensure df_proc is available from previous steps before generating charts
    if 'df_proc' in locals():
        cert_cols = find_monthly_columns(df_proc, "mto_certificado_")
        if cert_cols:
            cert_fig, cert_series = generate_monthly_series_chart(df_proc, cert_cols, "Certificado Mensual", "Certificado (S/)", 'lightcoral')
        else:
            cert_fig = None

        comp_cols = find_monthly_columns(df_proc, "mto_compro_")
        if comp_cols:
            comp_fig, comp_series = generate_monthly_series_chart(df_proc, comp_cols, "Comprometido Mensual", "Comprometido (S/)", 'cornflowerblue')
        else:
            comp_fig = None

        dev_cols = find_monthly_columns(df_proc, "mto_devenga_")
        if dev_cols:
            # Re-calculate dev_series based on the current filter for the monthly chart
            if 'selected_category' in locals() and selected_category != "Todos":
                df_dev_filtered_for_chart = df_proc[df_proc[group_col] == selected_category].copy()
            else:
                df_dev_filtered_for_chart = df_proc.copy()

            month_map_dev = {f"mto_devenga_{i:02d}": i for i in range(1, 13)}
            dev_series_for_chart = df_dev_filtered_for_chart[dev_cols].sum().reset_index()
            dev_series_for_chart.columns = ["col", "monto"]
            dev_series_for_chart["mes"] = dev_series_for_chart["col"].map(month_map_dev)
            dev_series_for_chart = dev_series_for_chart.sort_values("mes")

            dev_fig, dev_series = generate_monthly_series_chart(df_dev_filtered_for_chart, dev_cols, "Devengado Mensual", "Devengado (S/)","mediumseagreen")
        else:
            dev_fig = None
            dev_series = pd.DataFrame(columns=["mes", "monto"]) # Ensure dev_series is defined even if no dev_cols


        # =========================
        # Execution Projection Chart
        # =========================
        st.subheader("Proyección de Ejecución Acumulada")
        # Use the dev_series calculated for the monthly chart for projection
        # Ensure total_pim is available
        total_pim = float(df_proc.get("mto_pim", 0).sum())
        projection_fig = generate_projection_chart(dev_series, current_month, projection_method, total_pim)


        # =========================
        # Download Button (using modified function)
        # =========================
        monthly_charts_figs = {}
        if cert_fig: monthly_charts_figs["Grafico Certificado"] = cert_fig
        if comp_fig: monthly_charts_figs["Grafico Comprometido"] = comp_fig
        if dev_fig: monthly_charts_figs["Grafico Devengado"] = dev_fig


        # Ensure pivot and consolidado are defined before passing to download function
        # (They are defined in the sections above)
        if 'pivot' in locals() and 'consolidado' in locals():
            buf_modified = to_excel_download_modified(
                df_proc=df_proc,
                pivot=pivot,
                consolidado=consolidado,
                monthly_charts_figs=monthly_charts_figs,
                projection_fig=projection_fig,
                group_col=group_col,
                riesgo_umbral=riesgo_umbral
            )
            st.download_button(
                "Descargar Excel (Gráficos y Resumen Ejecutivo)",
                data=buf_modified,
                file_name="siaf_analisis_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
             st.warning("Run the previous steps to generate dataframes before downloading.")

    else:
        st.warning("Please upload and process the data first.")


# If uploaded is None, the initial message is shown and st.stop() is called.