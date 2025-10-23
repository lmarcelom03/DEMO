# -*- coding: utf-8 -*-

import base64
import io
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

EXCEL_SOURCE_DIR = Path(__file__).parent / "data" / "siaf"
EXCEL_SOURCE_DIR.mkdir(parents=True, exist_ok=True)


def _list_excel_candidates(folder: Path) -> List[Path]:
    return sorted(
        [p for p in folder.glob("*.xlsx") if p.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )


EXCEL_CANDIDATES = _list_excel_candidates(EXCEL_SOURCE_DIR)
LATEST_EXCEL = EXCEL_CANDIDATES[0] if EXCEL_CANDIDATES else None

PRIMARY_COLOR = "#c62828"
SECONDARY_COLOR = "#fbe9e7"
ACCENT_COLOR = "#0f4c81"

try:
    import xlsxwriter  # type: ignore
except ModuleNotFoundError:
    XLSXWRITER_AVAILABLE = False
else:
    XLSXWRITER_AVAILABLE = True

try:
    import openpyxl  # type: ignore # noqa: F401
except ModuleNotFoundError:
    OPENPYXL_AVAILABLE = False
else:
    OPENPYXL_AVAILABLE = True

LOGO_BASE64 = (
    "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"
)
LOGO_IMAGE = base64.b64decode(LOGO_BASE64)

APP_CSS = f"""
<style>
:root {{
    --primary-color: {PRIMARY_COLOR};
    --accent-color: {ACCENT_COLOR};
    --secondary-color: {SECONDARY_COLOR};
}}

[data-testid="stAppViewContainer"] {{
    background: linear-gradient(135deg, var(--secondary-color) 0%, #ffffff 55%);
}}

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, rgba(198,40,40,0.12), rgba(198,40,40,0));
}}

.app-title {{
    font-size: 2.4rem;
    font-weight: 700;
    margin-bottom: 0.15rem;
    color: var(--primary-color);
}}

.app-subtitle {{
    color: var(--accent-color);
    font-size: 1.1rem;
    margin-top: 0;
    margin-bottom: 0.6rem;
}}

.app-description {{
    color: #4a4a4a;
    font-size: 1.0rem;
    line-height: 1.55rem;
}}

.stTabs [data-baseweb="tab"] {{
    color: var(--accent-color);
    font-weight: 600;
}}

.stTabs [data-baseweb="tab"]:hover {{
    color: var(--primary-color);
    background-color: rgba(198, 40, 40, 0.08);
}}

.stTabs [data-baseweb="tab"][aria-selected="true"] {{
    color: var(--primary-color);
    border-bottom: 3px solid var(--primary-color);
}}

[data-testid="stMetricValue"] {{
    color: var(--primary-color);
}}

[data-testid="stMetricLabel"] {{
    color: #5c5c5c;
}}

.stRadio > div {{
    background-color: rgba(15,76,129,0.07);
    border-radius: 999px;
    padding: 0.35rem 0.75rem;
}}

.stRadio [data-baseweb="radio"] label span {{
    font-weight: 600;
    color: var(--accent-color);
}}

.stRadio [data-baseweb="radio"] input:checked + span {{
    color: var(--primary-color);
}}

.stButton>button {{
    background-color: var(--primary-color);
    color: #ffffff;
    font-weight: 600;
    border: none;
    box-shadow: 0 6px 16px rgba(198, 40, 40, 0.25);
}}

.stButton>button:hover {{
    background-color: #a12020;
}}
</style>
"""

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="SIAF Dashboard - Peru Compras", layout="wide")
st.markdown(APP_CSS, unsafe_allow_html=True)

header_col_logo, header_col_text = st.columns([1, 4])
with header_col_logo:
    st.image(LOGO_IMAGE, width=120)
with header_col_text:
    st.markdown("<h1 class='app-title'>SIAF Dashboard - Perú Compras</h1>", unsafe_allow_html=True)
    st.markdown(
        "<p class='app-subtitle'>Seguimiento diario del avance de ejecución presupuestal</p>",
        unsafe_allow_html=True,
    )

st.markdown(
    "<p class='app-description'>El dashboard toma automáticamente el <strong>Excel SIAF</strong> más reciente de la carpeta "
    "<code>data/siaf</code> para analizar <strong>PIA, PIM, Certificado, Comprometido, Devengado, Saldo PIM y % de avance</strong>. "
    "La aplicación asegura la lectura completa hasta CI, construye clasificadores jerárquicos estandarizados y ofrece vistas dinámicas con descargas.</p>",
    unsafe_allow_html=True,
)

# =========================
# Sidebar / parámetros
# =========================
selected_excel_path = LATEST_EXCEL

with st.sidebar:
    st.image(LOGO_IMAGE, width=140)
    st.markdown("<h3 style='color: var(--primary-color); margin-top: 0.5rem;'>Panel de control</h3>", unsafe_allow_html=True)
    st.header("Origen de datos")
    st.caption(
        "Coloca los archivos <code>.xlsx</code> en <code>data/siaf</code>. El dashboard usa el más reciente automáticamente."
    )
    if not EXCEL_CANDIDATES:
        st.error("No se encontraron archivos .xlsx en data/siaf. Añade uno y vuelve a actualizar.")
    else:
        label_to_path = {}
        option_labels = []
        for path in EXCEL_CANDIDATES:
            updated = datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            label = f"{path.name} · {updated}"
            label_to_path[label] = path
            option_labels.append(label)
        selected_label = st.selectbox(
            "Selecciona el archivo SIAF",
            options=option_labels,
            index=0,
            help="Los archivos están ordenados del más reciente al más antiguo.",
        )
        selected_excel_path = label_to_path[selected_label]
        st.success(f"Usando: {selected_excel_path.name}")
    st.markdown("---")
    st.header("Parámetros de lectura")
    usecols = st.text_input(
        "Rango de columnas (Excel)",
        "A:DV",
        help="Lectura fija para asegurar columnas CI–EC y programación mensual",
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

if selected_excel_path is None:
    st.error("No hay archivos disponibles en data/siaf. Añade un Excel y vuelve a ejecutar el dashboard.")
    st.stop()

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
    "programado",
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


def _flatten_headers(columns):
    """Normaliza encabezados (incluyendo multinivel) en snake_case en minúsculas."""

    flattened: List[str] = []
    seen_counts: Dict[str, int] = {}

    for col in columns:
        if isinstance(col, tuple):
            parts: List[str] = []
            for level in col:
                if level is None or (isinstance(level, float) and np.isnan(level)):
                    continue
                text = str(level).strip()
                if not text or text.lower() == "nan":
                    continue
                parts.append(text)
            label = " ".join(parts)
        else:
            if col is None or (isinstance(col, float) and np.isnan(col)):
                label = ""
            else:
                label = str(col).strip()

        if not label:
            label = "col"

        normalized = re.sub(r"\s+", "_", label)
        normalized = normalized.replace("__", "_")
        normalized = normalized.strip("_")
        normalized = normalized.lower() or "col"

        count = seen_counts.get(normalized, 0)
        seen_counts[normalized] = count + 1
        if count:
            normalized = f"{normalized}_{count+1}"

        flattened.append(normalized)

    return flattened


def load_data(excel_bytes, usecols, sheet_name, header_row_excel, autodetect=True):
    xls = pd.ExcelFile(excel_bytes)
    if autodetect:
        s, hdr = autodetect_sheet_and_header(xls, excel_bytes, usecols, sheet_name, header_row_excel)
    else:
        hdr = header_row_excel - 1
        s = sheet_name if sheet_name else xls.sheet_names[0]

    multi_header_df = None
    try:
        multi_header_df = pd.read_excel(
            excel_bytes,
            sheet_name=s,
            header=[hdr, hdr + 1],
            usecols=usecols,
        )
    except Exception:
        pass

    if multi_header_df is not None:
        df = multi_header_df
    else:
        df = pd.read_excel(excel_bytes, sheet_name=s, header=hdr, usecols=usecols)

    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = _flatten_headers(df.columns)
    return df, s

# =========================
# Cálculos CI–EC
# =========================
def find_monthly_columns(df, prefix):
    return [f"{prefix}{i:02d}" for i in range(1, 13) if f"{prefix}{i:02d}" in df.columns]


MONTH_NAME_ALIASES = {
    1: ("1", "01", "ene", "enero", "jan", "january"),
    2: ("2", "02", "feb", "febrero", "febr", "february"),
    3: ("3", "03", "mar", "marzo", "march"),
    4: ("4", "04", "abr", "abril", "april"),
    5: ("5", "05", "may", "mayo"),
    6: ("6", "06", "jun", "junio", "june"),
    7: ("7", "07", "jul", "julio", "july"),
    8: ("8", "08", "ago", "agosto", "aug", "august"),
    9: ("9", "09", "set", "sept", "septiembre", "sep", "september"),
    10: ("10", "oct", "octubre", "october"),
    11: ("11", "nov", "noviembre", "november"),
    12: ("12", "dic", "diciembre", "dec", "december"),
}
MONTH_NAME_LABELS = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Setiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}
PROGRAM_MATCH_TOKENS = ("prog", "program", "calendario", "cronograma")
PROGRAM_EXCLUDE_TOKENS = ("mto", "pia", "pim", "certificado", "compro", "girado", "devenga")


def _normalize_label(label) -> str:
    if label is None or (isinstance(label, float) and np.isnan(label)):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(label).lower())


def detect_programado_columns(df: pd.DataFrame) -> Dict[int, str]:
    """Infer the monthly programming columns (1-12) from the dataframe headers."""

    month_candidates: Dict[int, List[Tuple[str, bool]]] = {i: [] for i in range(1, 13)}
    fallback: List[str] = []

    for col in df.columns:
        series = df[col]
        if not pd.api.types.is_numeric_dtype(series):
            continue
        normalized = _normalize_label(col)
        if not normalized:
            continue
        contains_program = any(token in normalized for token in PROGRAM_MATCH_TOKENS)
        month_id = None
        for idx, aliases in MONTH_NAME_ALIASES.items():
            if any(alias in normalized for alias in aliases):
                month_id = idx
                break
        if month_id is not None:
            month_candidates[month_id].append((col, contains_program))
        elif contains_program:
            fallback.append(col)

    mapping: Dict[int, str] = {}
    for month_id, options in month_candidates.items():
        if not options:
            continue
        options_sorted = sorted(options, key=lambda item: (not item[1], len(str(item[0]))))
        mapping[month_id] = options_sorted[0][0]

    if len(mapping) < 12:
        numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
        ordered_candidates: List[str] = []
        for col in numeric_columns:
            if col in mapping.values():
                continue
            normalized = _normalize_label(col)
            if any(token in normalized for token in PROGRAM_EXCLUDE_TOKENS):
                continue
            ordered_candidates.append(col)
        for col in fallback:
            if col not in ordered_candidates and col not in mapping.values():
                ordered_candidates.append(col)

        for month_id in range(1, 13):
            if month_id in mapping:
                continue
            if not ordered_candidates:
                break
            mapping[month_id] = ordered_candidates.pop(0)

    return mapping


def attach_programado_metrics(df: pd.DataFrame, month: int):
    """Attach programado_mes and avance_programado_% columns based on detected schedule."""

    df = df.copy()
    month_map = detect_programado_columns(df)
    month_key = int(month) if pd.notna(month) else None
    source_col = month_map.get(month_key) if month_key in month_map else None

    if source_col and source_col in df.columns:
        program_series = pd.to_numeric(df[source_col], errors="coerce").fillna(0.0)
    else:
        program_series = pd.Series(0.0, index=df.index, dtype=float)
        source_col = None

    df["programado_mes"] = program_series.astype(float)
    devengado_mes = pd.to_numeric(df.get("devengado_mes", 0.0), errors="coerce").fillna(0.0).astype(float)
    df["devengado_mes"] = devengado_mes

    program_array = df["programado_mes"].to_numpy(dtype=float, copy=True)
    dev_array = devengado_mes.to_numpy(dtype=float, copy=True)
    ratio = np.zeros_like(program_array)
    np.divide(dev_array, program_array, out=ratio, where=program_array > 0)
    df["avance_programado_%"] = ratio * 100.0

    return df, month_map, source_col


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
    if "devengado_mes" in df.columns:
        cols.append("devengado_mes")
    if "programado_mes" in df.columns:
        cols.append("programado_mes")

    if "devengado" not in df.columns and dev_cols:
        df = df.copy()
        df["devengado"] = df[dev_cols].sum(axis=1)

    g = df.groupby(group_col, dropna=False)[cols].sum().reset_index()

    if "mto_pim" in g.columns and "devengado" in g.columns:
        g["saldo_pim"] = g["mto_pim"] - g["devengado"]
        g["avance_%"] = np.where(g["mto_pim"] > 0, g["devengado"] / g["mto_pim"] * 100.0, 0.0)
    if "mto_pim" in g.columns and "devengado_mes" in g.columns:
        g["avance_mes_%"] = np.where(g["mto_pim"] > 0, g["devengado_mes"] / g["mto_pim"] * 100.0, 0.0)
    if "programado_mes" in g.columns and "devengado_mes" in g.columns:
        g["avance_programado_%"] = np.where(
            g["programado_mes"] > 0,
            g["devengado_mes"] / g["programado_mes"] * 100.0,
            0.0,
        )

    return g


def _attach_pivot_table(
    workbook_buffer: io.BytesIO,
    source_sheet: str,
    target_sheet: str,
    table_name: str,
    row_fields: List[str],
    value_fields: List[Dict[str, object]],
):
    """Add an Excel pivot table using openpyxl after the workbook is written by XlsxWriter."""

    try:
        from openpyxl import load_workbook
        from openpyxl.pivot.cache import CacheDefinition, CacheField, CacheSource, WorksheetSource, SharedItems
        from openpyxl.pivot.table import (
            DataField,
            Location,
            PivotField,
            PivotTableStyle,
            RowColField,
            RowColItem,
            TableDefinition,
        )
        from openpyxl.utils import get_column_letter
    except Exception:
        return workbook_buffer

    workbook_buffer.seek(0)
    try:
        wb = load_workbook(workbook_buffer)
    except Exception:
        workbook_buffer.seek(0)
        return workbook_buffer

    if source_sheet not in wb.sheetnames:
        workbook_buffer.seek(0)
        return workbook_buffer

    ws_source = wb[source_sheet]
    max_row = ws_source.max_row
    max_col = ws_source.max_column

    if max_row <= 1 or max_col == 0:
        workbook_buffer.seek(0)
        return workbook_buffer

    headers = list(next(ws_source.iter_rows(min_row=1, max_row=1, values_only=True)))
    header_index = {name: idx for idx, name in enumerate(headers) if name}

    if not all(field in header_index for field in row_fields):
        workbook_buffer.seek(0)
        return workbook_buffer

    resolved_values = [vf for vf in value_fields if vf.get("field") in header_index]
    if not resolved_values:
        workbook_buffer.seek(0)
        return workbook_buffer

    if target_sheet in wb.sheetnames:
        del wb[target_sheet]
    ws_pivot = wb.create_sheet(target_sheet)

    data_ref = f"'{source_sheet}'!$A$1:${get_column_letter(max_col)}${max_row}"

    cache_fields = []
    for idx, header in enumerate(headers):
        if header is None:
            continue
        column_cells = list(
            ws_source.iter_cols(
                min_col=idx + 1,
                max_col=idx + 1,
                min_row=2,
                max_row=max_row,
                values_only=True,
            )
        )
        column_values = []
        if column_cells:
            column_values = [cell for cell in column_cells[0] if cell is not None]
        contains_number = any(isinstance(v, (int, float)) for v in column_values)
        contains_string = any(isinstance(v, str) for v in column_values)
        shared_items = SharedItems(
            count=len(column_values),
            containsNumber=contains_number or None,
            containsString=contains_string or None,
            containsBlank=True,
        )
        cache_fields.append(CacheField(name=str(header), sharedItems=shared_items))

    cache = CacheDefinition(
        cacheSource=CacheSource(
            type="worksheet",
            worksheetSource=WorksheetSource(ref=data_ref, sheet=source_sheet),
        ),
        recordCount=max_row - 1,
        cacheFields=tuple(cache_fields),
    )

    cache_id = len(wb._pivots) + 1 or 1
    cache.cacheId = cache_id
    cache._id = cache_id

    pivot_fields = []
    value_field_names = {vf["field"] for vf in resolved_values}
    row_indexes = [header_index[field] for field in row_fields]

    for idx, header in enumerate(headers):
        if header is None:
            continue
        pf = PivotField(name=str(header))
        if idx in row_indexes:
            pf.axis = "axisRow"
            pf.defaultSubtotal = True
        elif header in value_field_names:
            pf.dataField = True
        else:
            pf.defaultSubtotal = False
        pivot_fields.append(pf)

    row_fields_def = [RowColField(x=idx) for idx in row_indexes]
    row_items = [RowColItem(t="grand")]

    data_fields = []
    try:
        from openpyxl.styles.numbers import BUILTIN_FORMATS
    except Exception:
        BUILTIN_FORMATS = {}

    for value in resolved_values:
        field_name = value["field"]
        field_idx = header_index[field_name]
        subtotal = value.get("function", "sum")
        fmt_string = value.get("num_format")
        num_fmt_id = None
        if fmt_string:
            for fmt_key, fmt_val in BUILTIN_FORMATS.items():
                try:
                    matches = fmt_val == fmt_string
                except Exception:
                    matches = False
                if matches:
                    try:
                        num_fmt_id = int(fmt_key)
                    except Exception:
                        num_fmt_id = None
                    if num_fmt_id is not None:
                        break
        df = DataField(
            name=value.get("name", field_name),
            fld=field_idx,
            subtotal=subtotal,
            numFmtId=num_fmt_id,
        )
        data_fields.append(df)

    pivot_style = PivotTableStyle(name="PivotStyleMedium9", showRowHeaders=True, showColHeaders=True, showRowStripes=True)

    pivot = TableDefinition(
        name=table_name,
        cacheId=cache_id,
        dataOnRows=False,
        dataCaption="Valores",
        rowGrandTotals=True,
        colGrandTotals=True,
        location=Location(ref="A3", firstHeaderRow=3, firstDataRow=4, firstDataCol=1),
        pivotFields=tuple(pivot_fields),
        rowFields=tuple(row_fields_def),
        rowItems=tuple(row_items),
        dataFields=tuple(data_fields),
        pivotTableStyleInfo=pivot_style,
    )

    pivot.cache = cache
    ws_pivot.add_pivot(pivot)
    wb._pivots.append(pivot)

    updated_buffer = io.BytesIO()
    wb.save(updated_buffer)
    updated_buffer.seek(0)
    return updated_buffer


def to_excel_download(
    resumen,
    avance,
    proyeccion=None,
    ritmo=None,
    leaderboard=None,
    reporte_siaf=None,
    reporte_siaf_pivot_source=None,
):
    def _populate_workbook(writer: pd.ExcelWriter, use_xlsxwriter: bool):
        workbook = writer.book if use_xlsxwriter else None
        header_format = None
        currency_format = None
        percent_format = None
        if use_xlsxwriter and workbook is not None:
            header_format = workbook.add_format({"bold": True, "bg_color": "#c62828", "font_color": "#ffffff"})
            currency_format = workbook.add_format({"num_format": "#,##0.00"})
            percent_format = workbook.add_format({"num_format": "0.00%"})

        def _sanitize_table_name(name: str) -> str:
            clean = re.sub(r"[^0-9A-Za-z_]", "", name)[:20]
            return clean or "Tabla"

        def add_sheet_with_table(df: pd.DataFrame, sheet_name: str, add_chart: bool = True):
            if df is None or df.empty:
                return None

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]

            max_row, max_col = df.shape
            if use_xlsxwriter and workbook is not None:
                table_name = f"Tbl{_sanitize_table_name(sheet_name)}"
                worksheet.add_table(
                    0,
                    0,
                    max_row,
                    max_col - 1,
                    {
                        "name": table_name,
                        "style": "Table Style Medium 9",
                        "columns": [{"header": col} for col in df.columns],
                    },
                )

                worksheet.set_row(0, None, header_format)

                for col_idx, column_name in enumerate(df.columns):
                    if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                        fmt = percent_format if isinstance(column_name, str) and column_name.endswith("%") else currency_format
                        worksheet.set_column(col_idx, col_idx, None, fmt)

                if add_chart and max_row > 0 and max_col > 1:
                    chart = workbook.add_chart({"type": "column"})
                    categories = [sheet_name, 1, 0, max_row, 0]
                    for col_idx in range(1, max_col):
                        if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
                            chart.add_series(
                                {
                                    "name": [sheet_name, 0, col_idx],
                                    "categories": categories,
                                    "values": [sheet_name, 1, col_idx, max_row, col_idx],
                                }
                            )
                    chart.set_title({"name": sheet_name})
                    worksheet.insert_chart(1, max_col + 1, chart, {"x_scale": 1.1, "y_scale": 1.1})

            return worksheet

        add_sheet_with_table(resumen, "Resumen")
        add_sheet_with_table(avance, "Avance")

        if proyeccion is not None and not proyeccion.empty:
            add_sheet_with_table(proyeccion, "Proyeccion")
        if ritmo is not None and not ritmo.empty:
            add_sheet_with_table(ritmo, "Ritmo")
        if leaderboard is not None and not leaderboard.empty:
            add_sheet_with_table(leaderboard, "Leaderboard")

        pivot_source_sheet = None
        pivot_table_config = None
        if reporte_siaf is not None and not reporte_siaf.empty:
            add_sheet_with_table(reporte_siaf, "Reporte_SIAF")

        if reporte_siaf_pivot_source is not None and not reporte_siaf_pivot_source.empty:
            pivot_source_sheet = "Reporte_SIAF_Fuente"
            add_sheet_with_table(reporte_siaf_pivot_source, pivot_source_sheet, add_chart=False)
            pivot_table_config = {
                "rows": ["sec_func", "Generica", "clasificador_cod-concepto"],
                "values": [
                    {"field": "PIM", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "CERTIFICADO", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "COMPROMETIDO", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "DEVENGADO", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "DEVENGADO MES", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "PROGRAMADO MES", "function": "sum", "num_format": "#,##0.00"},
                    {"field": "Avance%", "function": "average", "num_format": "0.00%"},
                    {"field": "AvanceProgramado%", "function": "average", "num_format": "0.00%"},
                ],
            }

        return pivot_source_sheet, pivot_table_config

    engine_candidates = []
    missing_modules = set()
    if XLSXWRITER_AVAILABLE:
        engine_candidates.append("xlsxwriter")
    else:
        missing_modules.add("xlsxwriter")
    if OPENPYXL_AVAILABLE:
        engine_candidates.append("openpyxl")
    else:
        missing_modules.add("openpyxl")
    if not engine_candidates:
        missing_summary = ", ".join(sorted(missing_modules)) or "xlsxwriter, openpyxl"
        raise ModuleNotFoundError(
            f"No se encontró un motor de Excel disponible. Instala {missing_summary}.",
            name=missing_summary,
        )

    pivot_source_sheet = None
    pivot_table_config = None
    output = None
    engine_used = None
    last_exc = None

    for engine in engine_candidates:
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine=engine) as writer:
                pivot_source_sheet, pivot_table_config = _populate_workbook(writer, engine == "xlsxwriter")
            engine_used = engine
            break
        except ModuleNotFoundError as exc:
            missing_modules.add(getattr(exc, "name", engine))
            last_exc = exc
            continue

    if engine_used is None or output is None:
        missing_summary = ", ".join(sorted(missing_modules)) or "xlsxwriter, openpyxl"
        raise ModuleNotFoundError(
            f"No se encontró un motor de Excel disponible. Instala {missing_summary}.",
            name=missing_summary,
        ) from last_exc

    output.seek(0)
    if pivot_source_sheet and pivot_table_config:
        output = _attach_pivot_table(
            output,
            pivot_source_sheet,
            "Reporte_SIAF_Pivot",
            "PivotReporteSIAF",
            pivot_table_config["rows"],
            pivot_table_config["values"],
        )
    return output, engine_used

# =========================
# Carga del archivo
# =========================
try:
    df, used_sheet = load_data(
        selected_excel_path,
        usecols,
        sheet_name.strip() or None,
        int(header_row_excel),
        autodetect=detect_header,
    )
except Exception as e:
    st.error(f"No se pudo leer el archivo '{selected_excel_path.name}': {e}")
    st.stop()

st.success(
    f"Leída la hoja '{used_sheet}' del archivo '{selected_excel_path.name}' con {df.shape[0]} filas y {df.shape[1]} columnas."
)

if "sec_func" in df.columns:
    df["sec_func"] = df["sec_func"].apply(map_sec_func)

# =========================
# Filtros
# =========================
st.subheader("Filtros")
filter_cols = [c for c in df.columns if any(k in c for k in [
    "unidad_ejecutora","fuente_financ","generica","especifica_det","funcion",
    "programa_pptal","sec_func","departamento_meta","provincia_meta","area"
])]
filter_cols = [c for c in filter_cols if c not in {"subgenerica", "subgenerica_det"}]

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
df_proc, program_month_map, program_source_col = attach_programado_metrics(df_proc, current_month)

if program_month_map:
    month_label = MONTH_NAME_LABELS.get(int(current_month), f"Mes {int(current_month):02d}")
    if program_source_col:
        st.caption(
            f"Programación del mes {int(current_month):02d} ({month_label}) tomada de la columna "
            f"'{program_source_col}'."
        )
    else:
        st.caption(
            f"No se encontró columna de programación para el mes {int(current_month):02d} ({month_label}); se "
            "asumirá 0."
        )

    detected_pairs = [
        (MONTH_NAME_LABELS.get(month, f"Mes {month:02d}"), column)
        for month, column in sorted(program_month_map.items())
        if column
    ]
    if detected_pairs:
        items = "".join(
            f"<li><strong>{label}</strong>: <code>{column}</code></li>" for label, column in detected_pairs
        )
        st.markdown(
            f"<small>Columnas detectadas de programación mensual:<ul>{items}</ul></small>",
            unsafe_allow_html=True,
        )
else:
    st.caption("No se detectaron columnas de programación mensual en el archivo cargado.")

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
    "clasificador_cod","unidad_ejecutora","fuente_financ","generica","especifica_det",
    "funcion","programa_pptal","sec_func","area"
]]
_group_options = [c for c in _group_options if c not in {"subgenerica", "subgenerica_det"}]
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
    "devengado_mes","programado_mes","devengado","saldo_pim",
    "avance_%","avance_programado_%","riesgo_devolucion"
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
        "programado_mes",
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
    if "programado_mes" in consolidado.columns and "devengado_mes" in consolidado.columns:
        consolidado["avance_programado_%"] = np.where(
            consolidado["programado_mes"] > 0,
            consolidado["devengado_mes"] / consolidado["programado_mes"] * 100.0,
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
reporte_siaf_df = pd.DataFrame()
reporte_siaf_pivot_source = pd.DataFrame()
proyeccion_wide = pd.DataFrame()

# Navegación por apartados
(
    tab_resumen,
    tab_consol,
    tab_avance,
    tab_gestion,
    tab_reporte,
    tab_descarga,
) = st.tabs([
    "Resumen ejecutivo",
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
        vista_avance = st.radio(
            "Selecciona la vista",
            ("Gráfico", "Tabla"),
            horizontal=True,
            key="avance_view_mode",
            help="Alterna entre la visualización gráfica y la tabla resumen del devengado mensual.",
            label_visibility="collapsed",
        )

        if vista_avance == "Gráfico":
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
        else:
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
        st.info("No hay datos de PIM para calcular el ritmo requerido.")
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
        ritmo_df = round_numeric_for_reporting(pd.DataFrame(processes))
        if ritmo_df.empty:
            st.info("No hay información suficiente para calcular el ritmo requerido.")
        else:
            vista_ritmo = st.radio(
                "Selecciona la vista",
                ("Gráfico", "Tabla"),
                horizontal=True,
                key="ritmo_view_mode",
                help="Elige si deseas comparar el ritmo actual vs. necesario en gráfico o tabla.",
                label_visibility="collapsed",
            )

            if vista_ritmo == "Gráfico":
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
            else:
                fmt_ritmo = build_style_formatters(ritmo_df)
                ritmo_style = ritmo_df.style
                if fmt_ritmo:
                    ritmo_style = ritmo_style.format(fmt_ritmo)
                st.dataframe(ritmo_style, use_container_width=True)

    st.header("Top áreas con menor avance")
    if "sec_func" in df_view.columns and "mto_pim" in df_view.columns:
        agg_cols = ["mto_pim", "devengado", "devengado_mes", "programado_mes"]
        if "mto_certificado" in df_view.columns:
            agg_cols.insert(1, "mto_certificado")
        agg_sec = df_view.groupby("sec_func", dropna=False)[agg_cols].sum().reset_index()
        agg_sec = agg_sec[agg_sec["mto_pim"] > 0].copy()
        if agg_sec.empty:
            st.info("No hay áreas con PIM positivo para calcular el rendimiento.")
        else:
            agg_sec["avance_acum_%"] = np.where(agg_sec["mto_pim"] > 0, agg_sec["devengado"] / agg_sec["mto_pim"] * 100.0, 0.0)
            agg_sec["avance_mes_%"] = np.where(
                agg_sec["mto_pim"] > 0, agg_sec["devengado_mes"] / agg_sec["mto_pim"] * 100.0, 0.0,
            )
            agg_sec["avance_programado_%"] = np.where(
                agg_sec["programado_mes"] > 0,
                agg_sec["devengado_mes"] / agg_sec["programado_mes"] * 100.0,
                0.0,
            )
            agg_sec["rank_acum"] = agg_sec["avance_acum_%"].rank(method="dense", ascending=True).astype(int)
            agg_sec["rank_mes"] = agg_sec["avance_mes_%"].rank(method="dense", ascending=True).astype(int)

            max_top = int(agg_sec.shape[0])
            top_default = min(5, max_top) if max_top else 1
            top_n = st.slider("Número de áreas a mostrar", min_value=1, max_value=max_top or 1, value=top_default)

            leaderboard_df = (
                agg_sec.sort_values(["avance_acum_%", "avance_mes_%"], ascending=[True, True])
                .head(top_n)
                .copy()
            )
            display_cols = ["rank_acum", "rank_mes", "sec_func", "mto_pim"]
            if "mto_certificado" in agg_sec.columns:
                display_cols.append("mto_certificado")
            display_cols.extend([
                "devengado",
                "avance_acum_%",
                "devengado_mes",
                "programado_mes",
                "avance_mes_%",
                "avance_programado_%",
            ])
            leaderboard_df = leaderboard_df[display_cols]

            leaderboard_display = round_numeric_for_reporting(leaderboard_df)
            fmt_leader = build_style_formatters(leaderboard_display)
            highlight = lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else ""
            leader_style = leaderboard_display.style.applymap(
                highlight,
                subset=[c for c in ["avance_acum_%", "avance_mes_%", "avance_programado_%"] if c in leaderboard_display.columns],
            )
            if fmt_leader:
                leader_style = leader_style.format(fmt_leader)
            st.dataframe(leader_style, use_container_width=True)
    else:
        st.info("Se requieren las columnas sec_func y mto_pim para construir el ranking.")

with tab_reporte:
    st.header("Reporte SIAF por área, genérica y específica detalle")
    reporte_siaf_pivot_source = pd.DataFrame()
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
                "programado_mes",
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
                "DEVENGADO MES": "devengado_mes",
                "PROGRAMADO MES": "programado_mes",
            }
            for src in value_sources.values():
                if src not in reporte_base.columns:
                    reporte_base[src] = 0.0

            reporte_base = reporte_base[
                reporte_base["clasificador_cod_concepto"].fillna("").astype(str).str.strip() != ""
            ].copy()

            if not reporte_base.empty:
                def _safe_numeric(col_name):
                    if col_name in reporte_base.columns:
                        return reporte_base[col_name].fillna(0.0).astype(float)
                    return pd.Series(0.0, index=reporte_base.index, dtype=float)

                devengado_acum = _safe_numeric("devengado")
                devengado_mes_series = _safe_numeric("devengado_mes")
                programado_mes_series = _safe_numeric("programado_mes")
                pivot_source_df = pd.DataFrame(
                    {
                        "sec_func": reporte_base["sec_func"].fillna("").astype(str),
                        "Generica": reporte_base["generica"].fillna("").astype(str),
                        "clasificador_cod-concepto": reporte_base["clasificador_cod_concepto"].fillna("").astype(str),
                        "PIM": _safe_numeric("mto_pim"),
                        "CERTIFICADO": _safe_numeric("mto_certificado"),
                        "COMPROMETIDO": _safe_numeric("mto_compro_anual"),
                        "DEVENGADO": devengado_acum,
                        "DEVENGADO MES": devengado_mes_series,
                        "PROGRAMADO MES": programado_mes_series,
                    }
                )
                pivot_source_df["Avance%"] = np.where(
                    pivot_source_df["PIM"] > 0,
                    devengado_acum / pivot_source_df["PIM"],
                    0.0,
                )
                pivot_source_df["AvanceProgramado%"] = np.where(
                    pivot_source_df["PROGRAMADO MES"] > 0,
                    devengado_mes_series / pivot_source_df["PROGRAMADO MES"],
                    0.0,
                )
                reporte_siaf_pivot_source = pivot_source_df
            else:
                reporte_siaf_pivot_source = pd.DataFrame()

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
                reporte_siaf_df["% AVANCE DEV MES/PROG"] = np.where(
                    reporte_siaf_df["PROGRAMADO MES"].astype(float) > 0,
                    reporte_siaf_df["DEVENGADO MES"].astype(float)
                    / reporte_siaf_df["PROGRAMADO MES"].astype(float)
                    * 100.0,
                    0.0,
                )
                reporte_siaf_df = (
                    reporte_siaf_df.sort_values("orden", kind="stable")
                    .drop(columns=["orden", "nivel"], errors="ignore")
                )
                reporte_siaf_df = reporte_siaf_df[
                    [
                        "Centro de costo / Genérica de Gasto / Específica de Gasto",
                        "AVANCE DE EJECUCIÓN ACUMULADO",
                        "PIM",
                        "CERTIFICADO",
                        "COMPROMETIDO",
                        "DEVENGADO MES",
                        "PROGRAMADO MES",
                        "% AVANCE DEV MES/PROG",
                        "% AVANCE DEV /PIM",
                    ]
                ]
            else:
                reporte_siaf_df = pd.DataFrame(
                    columns=[
                        "Centro de costo / Genérica de Gasto / Específica de Gasto",
                        "AVANCE DE EJECUCIÓN ACUMULADO",
                        "PIM",
                        "CERTIFICADO",
                        "COMPROMETIDO",
                        "DEVENGADO MES",
                        "PROGRAMADO MES",
                        "% AVANCE DEV MES/PROG",
                        "% AVANCE DEV /PIM",
                    ]
                )

            reporte_display = round_numeric_for_reporting(reporte_siaf_df)
            fmt_reporte = build_style_formatters(reporte_display)
            reporte_style = reporte_display.style
            highlight_cols = [
                col
                for col in ["% AVANCE DEV /PIM", "% AVANCE DEV MES/PROG"]
                if col in reporte_display.columns
            ]
            if highlight_cols:
                reporte_style = reporte_style.applymap(
                    lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
                    subset=highlight_cols,
                )
            if fmt_reporte:
                reporte_style = reporte_style.format(fmt_reporte)
            st.dataframe(reporte_style, use_container_width=True)

with tab_descarga:
    st.header("Descarga de reportes")
    if not XLSXWRITER_AVAILABLE:
        st.warning(
            "No se encontró la librería `xlsxwriter`. El Excel se generará sin tablas ni gráficos embebidos."
        )
    excel_buffer = None
    excel_engine = None
    try:
        excel_buffer, excel_engine = to_excel_download(
            resumen=round_numeric_for_reporting(pivot.copy()),
            avance=round_numeric_for_reporting(avance_series.copy()),
            proyeccion=proyeccion_wide,
            ritmo=round_numeric_for_reporting(ritmo_df.copy()),
            leaderboard=round_numeric_for_reporting(leaderboard_df.copy()),
            reporte_siaf=round_numeric_for_reporting(reporte_siaf_df.copy()),
            reporte_siaf_pivot_source=reporte_siaf_pivot_source.copy(),
        )
    except ModuleNotFoundError as exc:
        missing = getattr(exc, "name", "xlsxwriter/openpyxl")
        st.error(
            "No se pudo generar el archivo de Excel porque faltan dependencias instaladas: "
            f"{missing}. Solicita al administrador que agregue el paquete correspondiente."
        )
    except Exception as exc:
        st.error(f"No se pudo generar el archivo de Excel: {exc}")
    else:
        if excel_engine == "openpyxl" and XLSXWRITER_AVAILABLE:
            st.info(
                "`xlsxwriter` no se pudo inicializar, se utilizó `openpyxl` como alternativa. "
                "Instala `xlsxwriter` para recuperar tablas y gráficos embebidos."
            )

    if excel_buffer is not None:
        st.download_button(
            "Descargar Excel (Resumen + Avance)",
            data=excel_buffer,
            file_name="siaf_resumen_avance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

