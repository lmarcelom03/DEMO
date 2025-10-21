
+# -*- coding: utf-8 -*-
+
 import base64
 import io
 import re
 import smtplib
 from email.message import EmailMessage
+from typing import Dict, List, Tuple
 
 import altair as alt
 import numpy as np
 import pandas as pd
 import streamlit as st
-from openpyxl import Workbook
-from openpyxl.chart import BarChart, Reference
-from openpyxl.utils import get_column_letter
-from openpyxl.utils.dataframe import dataframe_to_rows
-from openpyxl.worksheet.table import Table, TableStyleInfo
 
 PRIMARY_COLOR = "#c62828"
 SECONDARY_COLOR = "#fbe9e7"
 ACCENT_COLOR = "#0f4c81"
 
 LOGO_BASE64 = (
     "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"
@@ -143,52 +141,52 @@ header_col_logo, header_col_text = st.columns([1, 4])
 with header_col_logo:
     st.image(LOGO_IMAGE, width=120)
 with header_col_text:
     st.markdown("<h1 class='app-title'>SIAF Dashboard - Perú Compras</h1>", unsafe_allow_html=True)
     st.markdown(
         "<p class='app-subtitle'>Seguimiento diario del avance de ejecución presupuestal</p>",
         unsafe_allow_html=True,
     )
 
 st.markdown(
     "<p class='app-description'>Carga el <strong>Excel SIAF</strong> para analizar <strong>PIA, PIM, Certificado, Comprometido, Devengado, Saldo PIM y % de avance</strong>. "
     "La aplicación asegura la lectura completa hasta CI, construye clasificadores jerárquicos estandarizados y ofrece vistas dinámicas con descargas.</p>",
     unsafe_allow_html=True,
 )
 
 # =========================
 # Sidebar / parámetros
 # =========================
 with st.sidebar:
     st.image(LOGO_IMAGE, width=140)
     st.markdown("<h3 style='color: var(--primary-color); margin-top: 0.5rem;'>Panel de control</h3>", unsafe_allow_html=True)
     st.header("Parámetros de lectura")
     uploaded = st.file_uploader("Archivo SIAF (.xlsx)", type=["xlsx"])
     usecols = st.text_input(
         "Rango de columnas (Excel)",
-        "A:CH",
-        help="Lectura fija para asegurar columnas CI–EC",
+        "A:DV",
+        help="Lectura fija para asegurar columnas CI–EC y programación mensual",
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
@@ -225,50 +223,51 @@ def map_sec_func(value):
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
+    "programado",
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
@@ -299,98 +298,275 @@ def join_unique_nonempty(values, sep="\n"):
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
+        "programado": _safe_float(row.get("programado_mes", 0.0)),
+        "avance_programado": _safe_float(row.get("avance_programado_%", 0.0)),
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
 
 
+def _flatten_headers(columns):
+    """Normaliza encabezados (incluyendo multinivel) en snake_case en minúsculas."""
+
+    flattened: List[str] = []
+    seen_counts: Dict[str, int] = {}
+
+    for col in columns:
+        if isinstance(col, tuple):
+            parts: List[str] = []
+            for level in col:
+                if level is None or (isinstance(level, float) and np.isnan(level)):
+                    continue
+                text = str(level).strip()
+                if not text or text.lower() == "nan":
+                    continue
+                parts.append(text)
+            label = " ".join(parts)
+        else:
+            if col is None or (isinstance(col, float) and np.isnan(col)):
+                label = ""
+            else:
+                label = str(col).strip()
+
+        if not label:
+            label = "col"
+
+        normalized = re.sub(r"\s+", "_", label)
+        normalized = normalized.replace("__", "_")
+        normalized = normalized.strip("_")
+        normalized = normalized.lower() or "col"
+
+        count = seen_counts.get(normalized, 0)
+        seen_counts[normalized] = count + 1
+        if count:
+            normalized = f"{normalized}_{count+1}"
+
+        flattened.append(normalized)
+
+    return flattened
+
+
 def load_data(excel_bytes, usecols, sheet_name, header_row_excel, autodetect=True):
     xls = pd.ExcelFile(excel_bytes)
     if autodetect:
         s, hdr = autodetect_sheet_and_header(xls, excel_bytes, usecols, sheet_name, header_row_excel)
-        df = pd.read_excel(excel_bytes, sheet_name=s, header=hdr, usecols=usecols)
     else:
         hdr = header_row_excel - 1
         s = sheet_name if sheet_name else xls.sheet_names[0]
+
+    multi_header_df = None
+    try:
+        multi_header_df = pd.read_excel(
+            excel_bytes,
+            sheet_name=s,
+            header=[hdr, hdr + 1],
+            usecols=usecols,
+        )
+    except Exception:
+        pass
+
+    if multi_header_df is not None:
+        df = multi_header_df
+    else:
         df = pd.read_excel(excel_bytes, sheet_name=s, header=hdr, usecols=usecols)
 
     df = df.dropna(how="all").dropna(axis=1, how="all")
-    df.columns = [str(c).strip().lower() for c in df.columns]
+    df.columns = _flatten_headers(df.columns)
     return df, s
 
 # =========================
 # Cálculos CI–EC
 # =========================
 def find_monthly_columns(df, prefix):
     return [f"{prefix}{i:02d}" for i in range(1, 13) if f"{prefix}{i:02d}" in df.columns]
 
 
+MONTH_NAME_ALIASES = {
+    1: ("1", "01", "ene", "enero", "jan", "january"),
+    2: ("2", "02", "feb", "febrero", "febr", "february"),
+    3: ("3", "03", "mar", "marzo", "march"),
+    4: ("4", "04", "abr", "abril", "april"),
+    5: ("5", "05", "may", "mayo"),
+    6: ("6", "06", "jun", "junio", "june"),
+    7: ("7", "07", "jul", "julio", "july"),
+    8: ("8", "08", "ago", "agosto", "aug", "august"),
+    9: ("9", "09", "set", "sept", "septiembre", "sep", "september"),
+    10: ("10", "oct", "octubre", "october"),
+    11: ("11", "nov", "noviembre", "november"),
+    12: ("12", "dic", "diciembre", "dec", "december"),
+}
+MONTH_NAME_LABELS = {
+    1: "Enero",
+    2: "Febrero",
+    3: "Marzo",
+    4: "Abril",
+    5: "Mayo",
+    6: "Junio",
+    7: "Julio",
+    8: "Agosto",
+    9: "Setiembre",
+    10: "Octubre",
+    11: "Noviembre",
+    12: "Diciembre",
+}
+PROGRAM_MATCH_TOKENS = ("prog", "program", "calendario", "cronograma")
+PROGRAM_EXCLUDE_TOKENS = ("mto", "pia", "pim", "certificado", "compro", "girado", "devenga")
+
+
+def _normalize_label(label) -> str:
+    if label is None or (isinstance(label, float) and np.isnan(label)):
+        return ""
+    return re.sub(r"[^a-z0-9]", "", str(label).lower())
+
+
+def detect_programado_columns(df: pd.DataFrame) -> Dict[int, str]:
+    """Infer the monthly programming columns (1-12) from the dataframe headers."""
+
+    month_candidates: Dict[int, List[Tuple[str, bool]]] = {i: [] for i in range(1, 13)}
+    fallback: List[str] = []
+
+    for col in df.columns:
+        series = df[col]
+        if not pd.api.types.is_numeric_dtype(series):
+            continue
+        normalized = _normalize_label(col)
+        if not normalized:
+            continue
+        contains_program = any(token in normalized for token in PROGRAM_MATCH_TOKENS)
+        month_id = None
+        for idx, aliases in MONTH_NAME_ALIASES.items():
+            if any(alias in normalized for alias in aliases):
+                month_id = idx
+                break
+        if month_id is not None:
+            month_candidates[month_id].append((col, contains_program))
+        elif contains_program:
+            fallback.append(col)
+
+    mapping: Dict[int, str] = {}
+    for month_id, options in month_candidates.items():
+        if not options:
+            continue
+        options_sorted = sorted(options, key=lambda item: (not item[1], len(str(item[0]))))
+        mapping[month_id] = options_sorted[0][0]
+
+    if len(mapping) < 12:
+        numeric_columns = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
+        ordered_candidates: List[str] = []
+        for col in numeric_columns:
+            if col in mapping.values():
+                continue
+            normalized = _normalize_label(col)
+            if any(token in normalized for token in PROGRAM_EXCLUDE_TOKENS):
+                continue
+            ordered_candidates.append(col)
+        for col in fallback:
+            if col not in ordered_candidates and col not in mapping.values():
+                ordered_candidates.append(col)
+
+        for month_id in range(1, 13):
+            if month_id in mapping:
+                continue
+            if not ordered_candidates:
+                break
+            mapping[month_id] = ordered_candidates.pop(0)
+
+    return mapping
+
+
+def attach_programado_metrics(df: pd.DataFrame, month: int):
+    """Attach programado_mes and avance_programado_% columns based on detected schedule."""
+
+    df = df.copy()
+    month_map = detect_programado_columns(df)
+    month_key = int(month) if pd.notna(month) else None
+    source_col = month_map.get(month_key) if month_key in month_map else None
+
+    if source_col and source_col in df.columns:
+        program_series = pd.to_numeric(df[source_col], errors="coerce").fillna(0.0)
+    else:
+        program_series = pd.Series(0.0, index=df.index, dtype=float)
+        source_col = None
+
+    df["programado_mes"] = program_series.astype(float)
+    devengado_mes = pd.to_numeric(df.get("devengado_mes", 0.0), errors="coerce").fillna(0.0).astype(float)
+    df["devengado_mes"] = devengado_mes
+
+    program_array = df["programado_mes"].to_numpy(dtype=float, copy=True)
+    dev_array = devengado_mes.to_numpy(dtype=float, copy=True)
+    ratio = np.zeros_like(program_array)
+    np.divide(dev_array, program_array, out=ratio, where=program_array > 0)
+    df["avance_programado_%"] = ratio * 100.0
+
+    return df, month_map, source_col
+
+
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
@@ -499,112 +675,365 @@ def build_classifier_columns(df):
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
+    if "devengado_mes" in df.columns:
+        cols.append("devengado_mes")
+    if "programado_mes" in df.columns:
+        cols.append("programado_mes")
 
     if "devengado" not in df.columns and dev_cols:
         df = df.copy()
         df["devengado"] = df[dev_cols].sum(axis=1)
 
     g = df.groupby(group_col, dropna=False)[cols].sum().reset_index()
 
     if "mto_pim" in g.columns and "devengado" in g.columns:
         g["saldo_pim"] = g["mto_pim"] - g["devengado"]
         g["avance_%"] = np.where(g["mto_pim"] > 0, g["devengado"] / g["mto_pim"] * 100.0, 0.0)
+    if "mto_pim" in g.columns and "devengado_mes" in g.columns:
+        g["avance_mes_%"] = np.where(g["mto_pim"] > 0, g["devengado_mes"] / g["mto_pim"] * 100.0, 0.0)
+    if "programado_mes" in g.columns and "devengado_mes" in g.columns:
+        g["avance_programado_%"] = np.where(
+            g["programado_mes"] > 0,
+            g["devengado_mes"] / g["programado_mes"] * 100.0,
+            0.0,
+        )
 
     return g
 
 
+def _attach_pivot_table(
+    workbook_buffer: io.BytesIO,
+    source_sheet: str,
+    target_sheet: str,
+    table_name: str,
+    row_fields: List[str],
+    value_fields: List[Dict[str, object]],
+):
+    """Add an Excel pivot table using openpyxl after the workbook is written by XlsxWriter."""
+
+    try:
+        from openpyxl import load_workbook
+        from openpyxl.pivot.cache import CacheDefinition, CacheField, CacheSource, WorksheetSource, SharedItems
+        from openpyxl.pivot.table import (
+            DataField,
+            Location,
+            PivotField,
+            PivotTableStyle,
+            RowColField,
+            RowColItem,
+            TableDefinition,
+        )
+        from openpyxl.utils import get_column_letter
+    except Exception:
+        return workbook_buffer
+
+    workbook_buffer.seek(0)
+    try:
+        wb = load_workbook(workbook_buffer)
+    except Exception:
+        workbook_buffer.seek(0)
+        return workbook_buffer
+
+    if source_sheet not in wb.sheetnames:
+        workbook_buffer.seek(0)
+        return workbook_buffer
+
+    ws_source = wb[source_sheet]
+    max_row = ws_source.max_row
+    max_col = ws_source.max_column
+
+    if max_row <= 1 or max_col == 0:
+        workbook_buffer.seek(0)
+        return workbook_buffer
+
+    headers = list(next(ws_source.iter_rows(min_row=1, max_row=1, values_only=True)))
+    header_index = {name: idx for idx, name in enumerate(headers) if name}
+
+    if not all(field in header_index for field in row_fields):
+        workbook_buffer.seek(0)
+        return workbook_buffer
+
+    resolved_values = [vf for vf in value_fields if vf.get("field") in header_index]
+    if not resolved_values:
+        workbook_buffer.seek(0)
+        return workbook_buffer
+
+    if target_sheet in wb.sheetnames:
+        del wb[target_sheet]
+    ws_pivot = wb.create_sheet(target_sheet)
+
+    data_ref = f"'{source_sheet}'!$A$1:${get_column_letter(max_col)}${max_row}"
+
+    cache_fields = []
+    for idx, header in enumerate(headers):
+        if header is None:
+            continue
+        column_cells = list(
+            ws_source.iter_cols(
+                min_col=idx + 1,
+                max_col=idx + 1,
+                min_row=2,
+                max_row=max_row,
+                values_only=True,
+            )
+        )
+        column_values = []
+        if column_cells:
+            column_values = [cell for cell in column_cells[0] if cell is not None]
+        contains_number = any(isinstance(v, (int, float)) for v in column_values)
+        contains_string = any(isinstance(v, str) for v in column_values)
+        shared_items = SharedItems(
+            count=len(column_values),
+            containsNumber=contains_number or None,
+            containsString=contains_string or None,
+            containsBlank=True,
+        )
+        cache_fields.append(CacheField(name=str(header), sharedItems=shared_items))
+
+    cache = CacheDefinition(
+        cacheSource=CacheSource(
+            type="worksheet",
+            worksheetSource=WorksheetSource(ref=data_ref, sheet=source_sheet),
+        ),
+        recordCount=max_row - 1,
+        cacheFields=tuple(cache_fields),
+    )
+
+    cache_id = len(wb._pivots) + 1 or 1
+    cache.cacheId = cache_id
+    cache._id = cache_id
+
+    pivot_fields = []
+    value_field_names = {vf["field"] for vf in resolved_values}
+    row_indexes = [header_index[field] for field in row_fields]
+
+    for idx, header in enumerate(headers):
+        if header is None:
+            continue
+        pf = PivotField(name=str(header))
+        if idx in row_indexes:
+            pf.axis = "axisRow"
+            pf.defaultSubtotal = True
+        elif header in value_field_names:
+            pf.dataField = True
+        else:
+            pf.defaultSubtotal = False
+        pivot_fields.append(pf)
+
+    row_fields_def = [RowColField(x=idx) for idx in row_indexes]
+    row_items = [RowColItem(t="grand")]
+
+    data_fields = []
+    try:
+        from openpyxl.styles.numbers import BUILTIN_FORMATS
+    except Exception:
+        BUILTIN_FORMATS = {}
+
+    for value in resolved_values:
+        field_name = value["field"]
+        field_idx = header_index[field_name]
+        subtotal = value.get("function", "sum")
+        fmt_string = value.get("num_format")
+        num_fmt_id = None
+        if fmt_string:
+            for fmt_key, fmt_val in BUILTIN_FORMATS.items():
+                try:
+                    matches = fmt_val == fmt_string
+                except Exception:
+                    matches = False
+                if matches:
+                    try:
+                        num_fmt_id = int(fmt_key)
+                    except Exception:
+                        num_fmt_id = None
+                    if num_fmt_id is not None:
+                        break
+        df = DataField(
+            name=value.get("name", field_name),
+            fld=field_idx,
+            subtotal=subtotal,
+            numFmtId=num_fmt_id,
+        )
+        data_fields.append(df)
+
+    pivot_style = PivotTableStyle(name="PivotStyleMedium9", showRowHeaders=True, showColHeaders=True, showRowStripes=True)
+
+    pivot = TableDefinition(
+        name=table_name,
+        cacheId=cache_id,
+        dataOnRows=False,
+        rowGrandTotals=True,
+        colGrandTotals=True,
+        location=Location(ref="A3", firstHeaderRow=3, firstDataRow=4, firstDataCol=1),
+        pivotFields=tuple(pivot_fields),
+        rowFields=tuple(row_fields_def),
+        rowItems=tuple(row_items),
+        dataFields=tuple(data_fields),
+        pivotTableStyleInfo=pivot_style,
+    )
+
+    pivot.cache = cache
+    ws_pivot.add_pivot(pivot)
+    wb._pivots.append(pivot)
+
+    updated_buffer = io.BytesIO()
+    wb.save(updated_buffer)
+    updated_buffer.seek(0)
+    return updated_buffer
+
+
 def to_excel_download(
     resumen,
     avance,
     proyeccion=None,
     ritmo=None,
     leaderboard=None,
     reporte_siaf=None,
+    reporte_siaf_pivot_source=None,
 ):
-    wb = Workbook()
-    wb.remove(wb.active)
-
-    def add_table_with_chart(df, sheet_name):
-        ws = wb.create_sheet(sheet_name)
-        for r in dataframe_to_rows(df, index=False, header=True):
-            ws.append(r)
-        if ws.max_row <= 1:
-            return
-        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
-        tbl = Table(displayName=f"Tbl{sheet_name[:20].replace(' ','_')}", ref=ref)
-        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
-        ws.add_table(tbl)
-
-        num_cols = [i + 2 for i, c in enumerate(df.columns[1:]) if pd.api.types.is_numeric_dtype(df[c])]
-        if num_cols:
-            chart = BarChart()
-            data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row, max_col=max(num_cols))
-            cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
-            chart.add_data(data, titles_from_data=True)
-            chart.set_categories(cats)
-            chart.title = sheet_name
-            chart.height = 7
-            chart.width = 15
-            ws.add_chart(chart, f"{get_column_letter(ws.max_column + 2)}2")
-
-    add_table_with_chart(resumen, "Resumen")
-    add_table_with_chart(avance, "Avance")
-    if proyeccion is not None and not proyeccion.empty:
-        add_table_with_chart(proyeccion, "Proyeccion")
-    if ritmo is not None and not ritmo.empty:
-        add_table_with_chart(ritmo, "Ritmo")
-    if leaderboard is not None and not leaderboard.empty:
-        add_table_with_chart(leaderboard, "Leaderboard")
-    if reporte_siaf is not None and not reporte_siaf.empty:
-        add_table_with_chart(reporte_siaf, "Reporte_SIAF")
-
     output = io.BytesIO()
-    wb.save(output)
+
+    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
+        workbook = writer.book
+        header_format = workbook.add_format({"bold": True, "bg_color": "#c62828", "font_color": "#ffffff"})
+        currency_format = workbook.add_format({"num_format": "#,##0.00"})
+        percent_format = workbook.add_format({"num_format": "0.00%"})
+
+        def _sanitize_table_name(name: str) -> str:
+            clean = re.sub(r"[^0-9A-Za-z_]", "", name)[:20]
+            return clean or "Tabla"
+
+        def add_sheet_with_table(df: pd.DataFrame, sheet_name: str, add_chart: bool = True):
+            if df is None or df.empty:
+                return None
+
+            df.to_excel(writer, sheet_name=sheet_name, index=False)
+            worksheet = writer.sheets[sheet_name]
+
+            max_row, max_col = df.shape
+            table_name = f"Tbl{_sanitize_table_name(sheet_name)}"
+            worksheet.add_table(
+                0,
+                0,
+                max_row,
+                max_col - 1,
+                {
+                    "name": table_name,
+                    "style": "Table Style Medium 9",
+                    "columns": [{"header": col} for col in df.columns],
+                },
+            )
+
+            worksheet.set_row(0, None, header_format)
+
+            for col_idx, column_name in enumerate(df.columns):
+                if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
+                    fmt = percent_format if isinstance(column_name, str) and column_name.endswith("%") else currency_format
+                    worksheet.set_column(col_idx, col_idx, None, fmt)
+
+            if add_chart and max_row > 0 and max_col > 1:
+                chart = workbook.add_chart({"type": "column"})
+                categories = [sheet_name, 1, 0, max_row, 0]
+                for col_idx in range(1, max_col):
+                    if pd.api.types.is_numeric_dtype(df.iloc[:, col_idx]):
+                        chart.add_series(
+                            {
+                                "name": [sheet_name, 0, col_idx],
+                                "categories": categories,
+                                "values": [sheet_name, 1, col_idx, max_row, col_idx],
+                            }
+                        )
+                chart.set_title({"name": sheet_name})
+                worksheet.insert_chart(1, max_col + 1, chart, {"x_scale": 1.1, "y_scale": 1.1})
+
+            return worksheet
+
+        add_sheet_with_table(resumen, "Resumen")
+        add_sheet_with_table(avance, "Avance")
+
+        if proyeccion is not None and not proyeccion.empty:
+            add_sheet_with_table(proyeccion, "Proyeccion")
+        if ritmo is not None and not ritmo.empty:
+            add_sheet_with_table(ritmo, "Ritmo")
+        if leaderboard is not None and not leaderboard.empty:
+            add_sheet_with_table(leaderboard, "Leaderboard")
+
+        pivot_source_sheet = None
+        pivot_table_config = None
+        if reporte_siaf is not None and not reporte_siaf.empty:
+            add_sheet_with_table(reporte_siaf, "Reporte_SIAF")
+
+        if reporte_siaf_pivot_source is not None and not reporte_siaf_pivot_source.empty:
+            pivot_source_sheet = "Reporte_SIAF_Fuente"
+            add_sheet_with_table(reporte_siaf_pivot_source, pivot_source_sheet, add_chart=False)
+            pivot_table_config = {
+                "rows": ["sec_func", "Generica", "clasificador_cod-concepto"],
+                "values": [
+                    {"field": "PIM", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "CERTIFICADO", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "COMPROMETIDO", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "DEVENGADO", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "DEVENGADO MES", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "PROGRAMADO MES", "function": "sum", "num_format": "#,##0.00"},
+                    {"field": "Avance%", "function": "average", "num_format": "0.00%"},
+                    {"field": "AvanceProgramado%", "function": "average", "num_format": "0.00%"},
+                ],
+            }
+
     output.seek(0)
+    if pivot_source_sheet and pivot_table_config:
+        output = _attach_pivot_table(
+            output,
+            pivot_source_sheet,
+            "Reporte_SIAF_Pivot",
+            "PivotReporteSIAF",
+            pivot_table_config["rows"],
+            pivot_table_config["values"],
+        )
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
@@ -612,194 +1041,239 @@ filter_cols = [c for c in df.columns if any(k in c for k in [
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
+df_proc, program_month_map, program_source_col = attach_programado_metrics(df_proc, current_month)
+
+if program_month_map:
+    month_label = MONTH_NAME_LABELS.get(int(current_month), f"Mes {int(current_month):02d}")
+    if program_source_col:
+        st.caption(
+            f"Programación del mes {int(current_month):02d} ({month_label}) tomada de la columna "
+            f"'{program_source_col}'."
+        )
+    else:
+        st.caption(
+            f"No se encontró columna de programación para el mes {int(current_month):02d} ({month_label}); se "
+            "asumirá 0."
+        )
+
+    detected_pairs = [
+        (MONTH_NAME_LABELS.get(month, f"Mes {month:02d}"), column)
+        for month, column in sorted(program_month_map.items())
+        if column
+    ]
+    if detected_pairs:
+        items = "".join(
+            f"<li><strong>{label}</strong>: <code>{column}</code></li>" for label, column in detected_pairs
+        )
+        st.markdown(
+            f"<small>Columnas detectadas de programación mensual:<ul>{items}</ul></small>",
+            unsafe_allow_html=True,
+        )
+else:
+    st.caption("No se detectaron columnas de programación mensual en el archivo cargado.")
 
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
-    "devengado_mes","devengado","saldo_pim","avance_%","riesgo_devolucion"
+    "devengado_mes","programado_mes","devengado","saldo_pim",
+    "avance_%","avance_programado_%","riesgo_devolucion"
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
+        "programado_mes",
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
+    if "programado_mes" in consolidado.columns and "devengado_mes" in consolidado.columns:
+        consolidado["avance_programado_%"] = np.where(
+            consolidado["programado_mes"] > 0,
+            consolidado["devengado_mes"] / consolidado["programado_mes"] * 100.0,
+            0.0,
+        )
 
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
+reporte_siaf_pivot_source = pd.DataFrame()
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
-    if "avance_%" in pivot_display.columns:
+    highlight_cols = [
+        col
+        for col in ["avance_%", "avance_mes_%", "avance_programado_%"]
+        if col in pivot_display.columns
+    ]
+    if highlight_cols:
         pivot_style = pivot_style.applymap(
             lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
-            subset=["avance_%"],
+            subset=highlight_cols,
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
-        if "avance_%" in df_ci_display.columns:
+        highlight_cols = [c for c in ["avance_%", "avance_programado_%"] if c in df_ci_display.columns]
+        if highlight_cols:
             ci_style = ci_style.applymap(
                 lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
-                subset=["avance_%"],
+                subset=highlight_cols,
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
@@ -884,125 +1358,143 @@ with tab_gestion:
             ]
 
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
-        agg_cols = ["mto_pim", "devengado", "devengado_mes"]
+        agg_cols = ["mto_pim", "devengado", "devengado_mes", "programado_mes"]
         if "mto_certificado" in df_view.columns:
             agg_cols.insert(1, "mto_certificado")
         agg_sec = df_view.groupby("sec_func", dropna=False)[agg_cols].sum().reset_index()
         if agg_sec.empty:
             st.info("No hay datos disponibles para calcular el rendimiento por área.")
         else:
             agg_sec["avance_acum_%"] = np.where(agg_sec["mto_pim"] > 0, agg_sec["devengado"] / agg_sec["mto_pim"] * 100.0, 0.0)
             agg_sec["avance_mes_%"] = np.where(
                 agg_sec["mto_pim"] > 0, agg_sec["devengado_mes"] / agg_sec["mto_pim"] * 100.0, 0.0,
             )
+            agg_sec["avance_programado_%"] = np.where(
+                agg_sec["programado_mes"] > 0,
+                agg_sec["devengado_mes"] / agg_sec["programado_mes"] * 100.0,
+                0.0,
+            )
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
-            display_cols.extend(["devengado", "avance_acum_%", "devengado_mes", "avance_mes_%"])
+            display_cols.extend([
+                "devengado",
+                "avance_acum_%",
+                "devengado_mes",
+                "programado_mes",
+                "avance_mes_%",
+                "avance_programado_%",
+            ])
             leaderboard_df = leaderboard_df[display_cols]
 
             leaderboard_display = round_numeric_for_reporting(leaderboard_df)
             fmt_leader = build_style_formatters(leaderboard_display)
             highlight = lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else ""
             leader_style = leaderboard_display.style.applymap(
                 highlight,
-                subset=["avance_acum_%", "avance_mes_%"],
+                subset=[c for c in ["avance_acum_%", "avance_mes_%", "avance_programado_%"] if c in leaderboard_display.columns],
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
+            "Programado del mes: S/ {programado:,.2f}\n"
+            "Avance vs programado: {avance_programado:.2f}%\n\n"
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
-            subset=[c for c in ["avance_acum_%", "avance_mes_%"] if c in alert_display.columns],
+            subset=[
+                c
+                for c in ["avance_acum_%", "avance_mes_%", "avance_programado_%"]
+                if c in alert_display.columns
+            ],
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
@@ -1060,116 +1552,155 @@ with tab_gestion:
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
+    reporte_siaf_pivot_source = pd.DataFrame()
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
+                "programado_mes",
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
-                "DEVENGADO": "devengado_mes",
+                "DEVENGADO MES": "devengado_mes",
+                "PROGRAMADO MES": "programado_mes",
             }
             for src in value_sources.values():
                 if src not in reporte_base.columns:
                     reporte_base[src] = 0.0
 
             reporte_base = reporte_base[
                 reporte_base["clasificador_cod_concepto"].fillna("").astype(str).str.strip() != ""
             ].copy()
 
+            if not reporte_base.empty:
+                def _safe_numeric(col_name):
+                    if col_name in reporte_base.columns:
+                        return reporte_base[col_name].fillna(0.0).astype(float)
+                    return pd.Series(0.0, index=reporte_base.index, dtype=float)
+
+                devengado_acum = _safe_numeric("devengado")
+                devengado_mes_series = _safe_numeric("devengado_mes")
+                programado_mes_series = _safe_numeric("programado_mes")
+                pivot_source_df = pd.DataFrame(
+                    {
+                        "sec_func": reporte_base["sec_func"].fillna("").astype(str),
+                        "Generica": reporte_base["generica"].fillna("").astype(str),
+                        "clasificador_cod-concepto": reporte_base["clasificador_cod_concepto"].fillna("").astype(str),
+                        "PIM": _safe_numeric("mto_pim"),
+                        "CERTIFICADO": _safe_numeric("mto_certificado"),
+                        "COMPROMETIDO": _safe_numeric("mto_compro_anual"),
+                        "DEVENGADO": devengado_acum,
+                        "DEVENGADO MES": devengado_mes_series,
+                        "PROGRAMADO MES": programado_mes_series,
+                    }
+                )
+                pivot_source_df["Avance%"] = np.where(
+                    pivot_source_df["PIM"] > 0,
+                    devengado_acum / pivot_source_df["PIM"],
+                    0.0,
+                )
+                pivot_source_df["AvanceProgramado%"] = np.where(
+                    pivot_source_df["PROGRAMADO MES"] > 0,
+                    devengado_mes_series / pivot_source_df["PROGRAMADO MES"],
+                    0.0,
+                )
+                reporte_siaf_pivot_source = pivot_source_df
+            else:
+                reporte_siaf_pivot_source = pd.DataFrame()
+
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
 
@@ -1215,81 +1746,98 @@ with tab_reporte:
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
+                reporte_siaf_df["% AVANCE DEV MES/PROG"] = np.where(
+                    reporte_siaf_df["PROGRAMADO MES"].astype(float) > 0,
+                    reporte_siaf_df["DEVENGADO MES"].astype(float)
+                    / reporte_siaf_df["PROGRAMADO MES"].astype(float)
+                    * 100.0,
+                    0.0,
+                )
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
-                        "DEVENGADO",
+                        "DEVENGADO MES",
+                        "PROGRAMADO MES",
+                        "% AVANCE DEV MES/PROG",
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
-                        "DEVENGADO",
+                        "DEVENGADO MES",
+                        "PROGRAMADO MES",
+                        "% AVANCE DEV MES/PROG",
                         "% AVANCE DEV /PIM",
                     ]
                 )
 
             reporte_display = round_numeric_for_reporting(reporte_siaf_df)
             fmt_reporte = build_style_formatters(reporte_display)
             reporte_style = reporte_display.style
-            if "% AVANCE DEV /PIM" in reporte_display.columns:
+            highlight_cols = [
+                col
+                for col in ["% AVANCE DEV /PIM", "% AVANCE DEV MES/PROG"]
+                if col in reporte_display.columns
+            ]
+            if highlight_cols:
                 reporte_style = reporte_style.applymap(
                     lambda v: "background-color: #ffcccc" if v < float(riesgo_umbral) else "",
-                    subset=["% AVANCE DEV /PIM"],
+                    subset=highlight_cols,
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
+        reporte_siaf_pivot_source=reporte_siaf_pivot_source.copy(),
     )
     st.download_button(
         "Descargar Excel (Resumen + Avance)",
         data=buf,
         file_name="siaf_resumen_avance.xlsx",
         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
     )
