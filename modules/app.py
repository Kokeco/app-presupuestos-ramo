# app.py
# Aplicación Web Streamlit - Forecast 5+7 No Lineal Dinámico
# Proyecto de gastos mineros: carga cualquier Excel con estructura de gastos/presupuesto similar.

from __future__ import annotations

import datetime
import io
import re
import unicodedata
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Forecast 5+7 No Lineal Dinámico",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_ALIASES = {
    "jan": "Jan", "january": "Jan", "ene": "Jan", "enero": "Jan", "01": "Jan", "1": "Jan",
    "feb": "Feb", "february": "Feb", "febrero": "Feb", "02": "Feb", "2": "Feb",
    "mar": "Mar", "march": "Mar", "marzo": "Mar", "03": "Mar", "3": "Mar",
    "apr": "Apr", "april": "Apr", "abr": "Apr", "abril": "Apr", "04": "Apr", "4": "Apr",
    "may": "May", "mayo": "May", "05": "May", "5": "May",
    "jun": "Jun", "june": "Jun", "junio": "Jun", "06": "Jun", "6": "Jun",
    "jul": "Jul", "july": "Jul", "julio": "Jul", "07": "Jul", "7": "Jul",
    "aug": "Aug", "august": "Aug", "ago": "Aug", "agosto": "Aug", "08": "Aug", "8": "Aug",
    "sep": "Sep", "sept": "Sep", "september": "Sep", "septiembre": "Sep", "09": "Sep", "9": "Sep",
    "oct": "Oct", "october": "Oct", "octubre": "Oct", "10": "Oct",
    "nov": "Nov", "november": "Nov", "noviembre": "Nov", "11": "Nov",
    "dec": "Dec", "december": "Dec", "dic": "Dec", "diciembre": "Dec", "12": "Dec",
}

DEFAULT_CONTEXT = {
    "Labor": {"sens": 0.35, "mom": 0.10, "curve": 0.04, "min": 0.88, "max": 1.18, "why": "gasto estructural y relativamente estable"},
    "Fuel": {"sens": 0.75, "mom": 0.35, "curve": 0.18, "min": 0.70, "max": 1.65, "why": "volátil por precio, producción y flota"},
    "Contractors": {"sens": 0.65, "mom": 0.28, "curve": 0.16, "min": 0.72, "max": 1.55, "why": "depende del avance operacional y contratos"},
    "Maintenance": {"sens": 0.55, "mom": 0.30, "curve": 0.22, "min": 0.65, "max": 1.75, "why": "asociado a mantenciones y detenciones"},
    "Power": {"sens": 0.45, "mom": 0.20, "curve": 0.08, "min": 0.80, "max": 1.35, "why": "consumo energético operacional"},
    "Spare Parts": {"sens": 0.60, "mom": 0.25, "curve": 0.16, "min": 0.70, "max": 1.60, "why": "repuestos asociados a continuidad operacional"},
    "S&C": {"sens": 0.50, "mom": 0.20, "curve": 0.10, "min": 0.78, "max": 1.40, "why": "servicios y consumibles"},
    "Rehandling": {"sens": 0.60, "mom": 0.25, "curve": 0.12, "min": 0.72, "max": 1.50, "why": "movimiento adicional de material"},
    "Other": {"sens": 0.50, "mom": 0.20, "curve": 0.10, "min": 0.75, "max": 1.45, "why": "categoría general"},
}

POSSIBLE_CLASSIF = ["classif", "clasificacion", "clasificación", "naturaleza", "nature", "categoria", "categoría", "tipo gasto", "expense type"]
POSSIBLE_CC = ["cc", "centro costo", "centro de costo", "cost center", "resp", "responsable", "ceco"]
POSSIBLE_BUDGET_FY = ["budget fy", "budget_fy", "fy25", "fy26", "fy27", "fy28", "fy29", "fy30", "budget total", "presupuesto fy", "presupuesto anual"]
POSSIBLE_FORECAST_FY = ["forecast fy", "forecast_fy", "forecast actual", "proyeccion fy", "proyección fy"]


# ─────────────────────────────────────────────
# FUNCIONES UTILITARIAS
# ─────────────────────────────────────────────

def norm_txt(x: object) -> str:
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def find_month_in_col(col: object) -> Optional[str]:
    raw = str(col).strip()
    n = norm_txt(raw)
    if n.startswith("budget_"):
        n = n.replace("budget_", "")
    tokens = re.split(r"[^a-zA-Z0-9]+", n)
    tokens = [t for t in tokens if t]
    for t in tokens:
        if t in MONTH_ALIASES:
            return MONTH_ALIASES[t]
    for alias, std in MONTH_ALIASES.items():
        if re.search(rf"(^|[^a-z0-9]){re.escape(alias)}([^a-z0-9]|$)", n):
            return std
        if re.match(rf"^{re.escape(alias)}\d{{2,4}}$", n):
            return std
    m = re.search(r"20\d{2}[-_/ ](0?[1-9]|1[0-2])", n)
    if m:
        return MONTH_ALIASES[m.group(1).zfill(2)]
    return None


def detect_month_columns(df: pd.DataFrame) -> Dict[str, str]:
    found = {}
    for col in df.columns:
        m = find_month_in_col(col)
        if m and m not in found:
            nc = norm_txt(col)
            if any(bad in nc for bad in ["ytd", "bytd", "fy", "total", "var"]):
                continue
            found[m] = col
    return {m: found[m] for m in MONTH_ORDER if m in found}


def find_best_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized = {norm_txt(c): c for c in df.columns}
    for cand in candidates:
        if norm_txt(cand) in normalized:
            return normalized[norm_txt(cand)]
    for col in df.columns:
        nc = norm_txt(col)
        if any(norm_txt(cand) in nc for cand in candidates):
            return col
    return None


def money(value: float) -> str:
    try:
        value = float(value)
    except Exception:
        value = 0.0
    if abs(value) >= 1_000_000:
        return f"US$ {value/1_000_000:,.1f} MM"
    if abs(value) >= 1_000:
        return f"US$ {value/1_000:,.1f} k"
    return f"US$ {value:,.0f}"


def safe_div(num: float, den: float, default: float = 1.0) -> float:
    if pd.isna(den) or abs(float(den)) < 1e-9:
        return default
    return float(num) / float(den)


def scenario_curve(n: int, intensity: float, steepness: float) -> np.ndarray:
    if n <= 0:
        return np.array([])
    x = np.linspace(-2, 2, n)
    sig = 1 / (1 + np.exp(-steepness * x))
    centered = sig - sig.mean()
    return 1 + intensity * centered


def context_key(value: object) -> str:
    s = norm_txt(value)
    if "fuel" in s or "combust" in s or "diesel" in s:
        return "Fuel"
    if "contract" in s or "contrat" in s or "tercer" in s:
        return "Contractors"
    if "maint" in s or "mant" in s:
        return "Maintenance"
    if "labor" in s or "remun" in s or "sueldo" in s or "salary" in s:
        return "Labor"
    if "power" in s or "energia" in s or "electric" in s:
        return "Power"
    if "spare" in s or "repuesto" in s:
        return "Spare Parts"
    if "rehandling" in s or "remanejo" in s:
        return "Rehandling"
    return "Other"


def clean_sheet_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols_norm = [norm_txt(c) for c in df.columns]
    if "resp" in cols_norm and any("jan" in c or "ene" in c for c in cols_norm):
        return df
    max_rows = min(10, len(df))
    for i in range(max_rows):
        row_values = [norm_txt(x) for x in df.iloc[i].tolist()]
        has_resp = "resp" in row_values
        has_cc = "cc" in row_values
        has_month = any(
            any(alias in cell for alias in ["jan", "ene", "feb", "mar", "apr", "abr", "may", "jun", "jul", "aug", "ago", "sep", "oct", "nov", "dec", "dic"])
            for cell in row_values
        )
        if has_resp and has_cc and has_month:
            new_cols = df.iloc[i].astype(str).tolist()
            df = df.iloc[i + 1:].copy()
            df.columns = new_cols
            df = df.dropna(how="all")
            df = df.loc[:, ~pd.Index(df.columns).astype(str).str.startswith("nan")]
            return df.reset_index(drop=True)
    return df


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def read_excel_sheets(uploaded_file) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for sheet in xls.sheet_names:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet)
        sheets[sheet] = clean_sheet_headers(raw)
    return sheets


# ─────────────────────────────────────────────
# PREPARACIÓN Y FORECAST
# ─────────────────────────────────────────────

def prepare_data(
    gastos_raw: pd.DataFrame,
    budget_raw: Optional[pd.DataFrame],
    gastos_month_cols: Dict[str, str],
    budget_month_cols: Dict[str, str],
    actual_months: List[str],
    forecast_months: List[str],
    classif_col: Optional[str],
    cc_col_gastos: Optional[str],
    cc_col_budget: Optional[str],
    budget_fy_col: Optional[str],
    extra_dims: List[str],
) -> pd.DataFrame:
    df = gastos_raw.copy()
    for month, col in gastos_month_cols.items():
        df[f"Actual_{month}"] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if classif_col and classif_col in df.columns:
        df["Naturaleza"] = df[classif_col].fillna("Other").astype(str)
    else:
        df["Naturaleza"] = "Other"
    df["Contexto_Mina"] = df["Naturaleza"].map(context_key)

    if cc_col_gastos and cc_col_gastos in df.columns:
        df["_CC_KEY_"] = df[cc_col_gastos].astype(str).str.replace(".0", "", regex=False).str.strip()
    else:
        df["_CC_KEY_"] = np.arange(len(df)).astype(str)

    for m in MONTH_ORDER:
        df[f"Budget_{m}"] = 0.0

    if budget_raw is not None and budget_month_cols:
        b = budget_raw.copy()
        if cc_col_budget and cc_col_budget in b.columns:
            b["_CC_KEY_"] = b[cc_col_budget].astype(str).str.replace(".0", "", regex=False).str.strip()
        else:
            b["_CC_KEY_"] = np.arange(len(b)).astype(str)
        agg_dict = {}
        for m, col in budget_month_cols.items():
            b[f"Budget_{m}"] = pd.to_numeric(b[col], errors="coerce").fillna(0)
            agg_dict[f"Budget_{m}"] = "sum"
        if budget_fy_col and budget_fy_col in b.columns:
            b["Budget_FY_Input"] = pd.to_numeric(b[budget_fy_col], errors="coerce").fillna(0)
            agg_dict["Budget_FY_Input"] = "sum"
        b_agg = b.groupby("_CC_KEY_", as_index=False).agg(agg_dict)
        df = df.merge(b_agg, on="_CC_KEY_", how="left", suffixes=("", "_from_budget"))
        for m in MONTH_ORDER:
            merged = f"Budget_{m}_from_budget"
            if merged in df.columns:
                df[f"Budget_{m}"] = pd.to_numeric(df[merged], errors="coerce").fillna(df[f"Budget_{m}"])
                df.drop(columns=[merged], inplace=True)
    else:
        for col in gastos_raw.columns:
            nc = norm_txt(col)
            if "budget" in nc or "presupuesto" in nc:
                m = find_month_in_col(col)
                if m:
                    df[f"Budget_{m}"] = pd.to_numeric(gastos_raw[col], errors="coerce").fillna(0)

    for m in MONTH_ORDER:
        if f"Budget_{m}" not in df.columns:
            df[f"Budget_{m}"] = 0.0
        df[f"Budget_{m}"] = pd.to_numeric(df[f"Budget_{m}"], errors="coerce").fillna(0)

    df["Actual_YTD"] = df[[f"Actual_{m}" for m in actual_months if f"Actual_{m}" in df.columns]].sum(axis=1)
    df["Budget_YTD"] = df[[f"Budget_{m}" for m in actual_months]].sum(axis=1)
    df["Budget_Remaining"] = df[[f"Budget_{m}" for m in forecast_months]].sum(axis=1)
    df["Budget_FY_Model"] = df[[f"Budget_{m}" for m in MONTH_ORDER]].sum(axis=1)
    if "Budget_FY_Input" in df.columns:
        df["Budget_FY_Model"] = df["Budget_FY_Input"].where(df["Budget_FY_Input"].abs() > 1e-9, df["Budget_FY_Model"])

    df["Budget_YTD"] = np.where(df["Budget_YTD"].abs() < 1e-9, df["Actual_YTD"], df["Budget_YTD"])
    return df


def calculate_forecast(
    df: pd.DataFrame,
    actual_months: List[str],
    forecast_months: List[str],
    sensitivity_mult: float,
    momentum_mult: float,
    scenario_mult: float,
    steepness: float,
) -> pd.DataFrame:
    out = df.copy()
    if len(actual_months) == 0 or len(forecast_months) == 0:
        raise ValueError("Debes tener al menos 1 mes real y 1 mes forecast.")

    recent = actual_months[-min(2, len(actual_months)):]
    early = actual_months[:min(3, len(actual_months))]
    out["Prom_Reciente"] = out[[f"Actual_{m}" for m in recent]].mean(axis=1)
    out["Prom_Inicial"] = out[[f"Actual_{m}" for m in early]].mean(axis=1)
    out["Factor_Ejecucion"] = [safe_div(a, b, 1.0) for a, b in zip(out["Actual_YTD"], out["Budget_YTD"])]
    out["Factor_Tendencia_Bruto"] = [safe_div(a, b, 1.0) for a, b in zip(out["Prom_Reciente"], out["Prom_Inicial"])]
    out["Factor_Ejecucion"] = out["Factor_Ejecucion"].clip(0.2, 2.8)
    out["Factor_Tendencia_Bruto"] = out["Factor_Tendencia_Bruto"].clip(0.35, 2.5)

    for m in forecast_months:
        out[f"Forecast_{m}"] = 0.0

    for idx, row in out.iterrows():
        params = DEFAULT_CONTEXT.get(row["Contexto_Mina"], DEFAULT_CONTEXT["Other"])
        sens = params["sens"] * sensitivity_mult
        mom = params["mom"] * momentum_mult
        curve_intensity = params["curve"] * scenario_mult
        min_f, max_f = params["min"], params["max"]
        base_factor = 1 + sens * (row["Factor_Ejecucion"] - 1) + mom * (row["Factor_Tendencia_Bruto"] - 1)
        base_factor = float(np.clip(base_factor, min_f, max_f))
        curve = scenario_curve(len(forecast_months), curve_intensity, steepness)
        for j, m in enumerate(forecast_months):
            budget = row.get(f"Budget_{m}", 0.0)
            fallback = row["Actual_YTD"] / max(len(actual_months), 1)
            base = budget if abs(budget) > 1e-9 else fallback
            out.at[idx, f"Forecast_{m}"] = base * base_factor * curve[j]

    out["Forecast_Remaining"] = out[[f"Forecast_{m}" for m in forecast_months]].sum(axis=1)
    out["Forecast_FY_Modelo"] = out["Actual_YTD"] + out["Forecast_Remaining"]
    out["Var_vs_Budget"] = out["Forecast_FY_Modelo"] - out["Budget_FY_Model"]
    out["Var_vs_Budget_%"] = [safe_div(v, b, 0.0) for v, b in zip(out["Var_vs_Budget"], out["Budget_FY_Model"])]
    out["Recomendacion"] = np.select(
        [out["Var_vs_Budget_%"] > 0.10, out["Var_vs_Budget_%"] < -0.10],
        ["Riesgo de sobreconsumo: activar plan de control y revisión contractual.", "Subejecución relevante: revisar reprogramación operacional o liberación de presupuesto."],
        default="Desviación controlada: mantener monitoreo mensual.",
    )
    out["Justificacion_Mina"] = out["Contexto_Mina"].map(lambda k: DEFAULT_CONTEXT.get(k, DEFAULT_CONTEXT["Other"])["why"])
    return out


def aggregate_for_charts(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if group_col not in df.columns:
        group_col = "Naturaleza"
    return df.groupby(group_col, dropna=False, as_index=False).agg(
        Actual_YTD=("Actual_YTD", "sum"),
        Budget_FY_Model=("Budget_FY_Model", "sum"),
        Forecast_FY_Modelo=("Forecast_FY_Modelo", "sum"),
        Var_vs_Budget=("Var_vs_Budget", "sum"),
    ).sort_values("Forecast_FY_Modelo", ascending=False)


# ─────────────────────────────────────────────
# SIMULACIÓN 5 AÑOS
# ─────────────────────────────────────────────

def calcular_estacionalidad_mensual(
    df: pd.DataFrame,
    actual_months: List[str],
    forecast_months: List[str],
) -> Dict[str, float]:
    """Calcula el % que representa cada mes sobre el total FY (Actual + Forecast).
    Se usa como patrón de estacionalidad para repartir Año_1 en 12 meses."""
    total_fy = 0.0
    montos_mes = {}
    for m in MONTH_ORDER:
        if m in actual_months:
            col = f"Actual_{m}"
        elif m in forecast_months:
            col = f"Forecast_{m}"
        else:
            col = None
        monto = float(df[col].sum()) if col and col in df.columns else 0.0
        montos_mes[m] = monto
        total_fy += monto
    if total_fy <= 1e-9:
        return {m: 1 / 12 for m in MONTH_ORDER}
    return {m: montos_mes[m] / total_fy for m in MONTH_ORDER}


def simular_5_anos(df_base: pd.DataFrame, inf_anual: float, crec_ops: float) -> pd.DataFrame:
    """Proyecta 5 años usando el Forecast actual como base y tasas de crecimiento compuestas."""
    df_5y = df_base.copy()

    def obtener_tasa(ctx):
        if ctx in ["Labor", "Other"]:
            return inf_anual / 100
        if ctx in ["Fuel", "Power", "Spare Parts", "Rehandling"]:
            return (inf_anual + crec_ops) / 100
        return (inf_anual + (crec_ops * 0.5)) / 100

    df_5y["Tasa_Crecimiento"] = df_5y["Contexto_Mina"].apply(obtener_tasa)
    df_5y["Año_0_Base"] = df_5y["Forecast_FY_Modelo"]

    # Fix: prev_col evita buscar "Año_0" que no existe
    prev_col = "Año_0_Base"
    for i in range(1, 6):
        new_col = f"Año_{i}"
        df_5y[new_col] = df_5y[prev_col] * (1 + df_5y["Tasa_Crecimiento"])
        prev_col = new_col

    return df_5y


# ─────────────────────────────────────────────
# EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────

def to_excel_bytes(df: pd.DataFrame, summary: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Forecast_Detalle")
        summary.to_excel(writer, index=False, sheet_name="Resumen")
        wb = writer.book
        money_fmt = wb.add_format({"num_format": '#,##0;[Red]-#,##0'})
        pct_fmt = wb.add_format({"num_format": '0.0%'})
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        for sheet_name in ["Forecast_Detalle", "Resumen"]:
            ws = writer.sheets[sheet_name]
            for c, col in enumerate((df if sheet_name == "Forecast_Detalle" else summary).columns):
                ws.write(0, c, col, header_fmt)
                width = min(max(len(str(col)) + 2, 12), 35)
                ws.set_column(c, c, width)
                if any(k in str(col).lower() for k in ["budget", "forecast", "actual", "var", "prom"]):
                    ws.set_column(c, c, width, money_fmt)
                if "%" in str(col):
                    ws.set_column(c, c, width, pct_fmt)
    return output.getvalue()


def to_excel_completo(
    df_forecast: pd.DataFrame,
    summary_forecast: pd.DataFrame,
    resumen_5y: pd.DataFrame,
    resumen_largo: pd.DataFrame,
    actual_months: List[str],
    forecast_months: List[str],
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        money_fmt  = wb.add_format({"num_format": '#,##0;[Red]-#,##0'})
        pct_fmt    = wb.add_format({"num_format": '0.0%'})
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        green_fmt  = wb.add_format({"bold": True, "bg_color": "#C6EFCE", "border": 1})
        red_fmt    = wb.add_format({"bold": True, "bg_color": "#FFC7CE", "border": 1})

        def write_sheet(df, sheet_name, money_keys=(), pct_keys=()):
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for c, col in enumerate(df.columns):
                ws.write(0, c, col, header_fmt)
                width = min(max(len(str(col)) + 2, 14), 38)
                if any(k in str(col).lower() for k in pct_keys):
                    ws.set_column(c, c, width, pct_fmt)
                elif any(k in str(col).lower() for k in money_keys):
                    ws.set_column(c, c, width, money_fmt)
                else:
                    ws.set_column(c, c, width)

        # Hoja 1: Resumen Forecast
        write_sheet(
            summary_forecast, "Resumen_Forecast",
            money_keys=["actual", "budget", "forecast", "var"],
            pct_keys=["%"],
        )

        # Hoja 2: Detalle Forecast fila a fila
        cols_detalle = (
            ["Naturaleza", "Contexto_Mina"]
            + [f"Actual_{m}"   for m in actual_months   if f"Actual_{m}"   in df_forecast.columns]
            + [f"Budget_{m}"   for m in actual_months   if f"Budget_{m}"   in df_forecast.columns]
            + [f"Forecast_{m}" for m in forecast_months if f"Forecast_{m}" in df_forecast.columns]
            + ["Actual_YTD", "Budget_FY_Model", "Forecast_FY_Modelo",
               "Var_vs_Budget", "Var_vs_Budget_%", "Recomendacion"]
        )
        cols_detalle = list(dict.fromkeys([c for c in cols_detalle if c in df_forecast.columns]))
        write_sheet(
            df_forecast[cols_detalle], "Detalle_Forecast",
            money_keys=["actual", "budget", "forecast", "var", "prom"],
            pct_keys=["%"],
        )

        # Colorear Var_vs_Budget_% en Detalle_Forecast
        ws2 = writer.sheets["Detalle_Forecast"]
        if "Var_vs_Budget_%" in cols_detalle:
            col_idx = cols_detalle.index("Var_vs_Budget_%")
            for row_idx, val in enumerate(df_forecast["Var_vs_Budget_%"], start=1):
                try:
                    v = float(val)
                    fmt = red_fmt if v > 0.10 else (green_fmt if v < -0.10 else pct_fmt)
                except Exception:
                    fmt = pct_fmt
                ws2.write(row_idx, col_idx, val, fmt)

        # Hoja 3: Simulación LRP 5 Años (matriz con meses de Año_1)
        write_sheet(
            resumen_5y, "Simulacion_5_Anos",
            money_keys=["fy", "jan", "feb", "mar", "apr", "may", "jun",
                        "jul", "aug", "sep", "oct", "nov", "dec"],
        )

        # Hoja 4: LRP formato largo (para pivots)
        write_sheet(
            resumen_largo, "LRP_Detalle_Largo",
            money_keys=["presupuesto"],
        )

        # Hoja 5: Parámetros del modelo
        params_data = {
            "Parámetro": [
                "Último mes real (YTD)", "Meses reales", "Meses forecast",
                "Sensibilidad ejecución YTD", "Peso tendencia reciente",
                "Intensidad curva no lineal", "Forma curva logística",
                "Inflación anual estimada (%)", "Crecimiento operacional anual (%)",
            ],
            "Valor": [
                cutoff_month, ", ".join(actual_months), ", ".join(forecast_months),
                sensitivity_mult, momentum_mult, scenario_mult, steepness,
                cagr_inf, cagr_ops,
            ],
        }
        pd.DataFrame(params_data).to_excel(writer, index=False, sheet_name="Parametros_Modelo")
        ws5 = writer.sheets["Parametros_Modelo"]
        ws5.set_column(0, 0, 35)
        ws5.set_column(1, 1, 45)
        for c, col in enumerate(["Parámetro", "Valor"]):
            ws5.write(0, c, col, header_fmt)

    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ═══════════════════════════════════════════════════════════════

st.title("⛏️ Forecast 5+7 No Lineal Dinámico")
st.caption("Aplicación web para presupuesto y gastos mineros: carga Excel, detecta meses, proyecta forecast y genera hallazgos ejecutivos.")

with st.expander("📌 Metodología del modelo", expanded=False):
    st.markdown(
        """
        **Forecast mensual futuro = Budget mensual × Factor de ejecución × Factor de tendencia × Curva no lineal**

        - **Factor de ejecución:** compara gasto real acumulado contra presupuesto acumulado.
        - **Factor de tendencia:** compara meses recientes contra meses iniciales.
        - **Curva no lineal:** distribuye el ajuste en los meses restantes con una forma logística, evitando una simple extrapolación lineal.
        - **Contexto minero:** los factores cambian según naturaleza del gasto: Fuel, Contractors, Maintenance, Labor, Power, Spare Parts u Other.
        """
    )

uploaded = st.sidebar.file_uploader("Sube tu Excel de gastos/presupuesto", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Sube un archivo Excel para comenzar. La app funciona con hojas similares a 'Gastos' y 'Budget', pero permite seleccionar otra estructura.")
    st.stop()

try:
    sheets = read_excel_sheets(uploaded)
except Exception as e:
    st.error(f"No pude leer el Excel: {e}")
    st.stop()

sheet_names = list(sheets.keys())
def_sheet_gastos = next((s for s in sheet_names if norm_txt(s) in ["gastos", "expenses", "costs"]), sheet_names[0])
def_sheet_budget = next((s for s in sheet_names if "budget" in norm_txt(s) or "presupuesto" in norm_txt(s)), sheet_names[0])

st.sidebar.subheader("1) Selección de hojas")
gastos_sheet = st.sidebar.selectbox("Hoja de gastos reales", sheet_names, index=sheet_names.index(def_sheet_gastos))
budget_sheet = st.sidebar.selectbox("Hoja de presupuesto", ["Sin hoja Budget / usar misma hoja"] + sheet_names, index=(sheet_names.index(def_sheet_budget) + 1 if def_sheet_budget in sheet_names else 0))

gastos_raw = sheets[gastos_sheet].copy()
budget_raw = None if budget_sheet.startswith("Sin hoja") else sheets[budget_sheet].copy()

gastos_month_cols = detect_month_columns(gastos_raw)
budget_month_cols = detect_month_columns(budget_raw) if budget_raw is not None else {}

if len(gastos_month_cols) < 2:
    st.error("No detecté suficientes columnas mensuales en la hoja de gastos. Revisa que existan columnas tipo Jan-25, Ene-25, 2025-01, etc.")
    st.write("Columnas detectadas:", list(gastos_raw.columns))
    st.stop()

available_months = [m for m in MONTH_ORDER if m in gastos_month_cols]

st.sidebar.subheader("2) Configuración Forecast")
cutoff_default = available_months[min(4, len(available_months)-2)] if len(available_months) >= 6 else available_months[-2]
cutoff_month = st.sidebar.selectbox("Último mes real / cierre YTD", available_months[:-1], index=available_months[:-1].index(cutoff_default) if cutoff_default in available_months[:-1] else 0)
cutoff_idx = MONTH_ORDER.index(cutoff_month)
actual_months = [m for m in MONTH_ORDER[:cutoff_idx+1] if m in gastos_month_cols]
forecast_months = [m for m in MONTH_ORDER[cutoff_idx+1:] if m in MONTH_ORDER]

st.sidebar.write(f"Meses reales: {', '.join(actual_months)}")
st.sidebar.write(f"Meses forecast: {', '.join(forecast_months)}")

st.sidebar.subheader("3) Columnas de negocio")
all_cols = [None] + list(gastos_raw.columns)
classif_guess = find_best_col(gastos_raw, POSSIBLE_CLASSIF)
cc_guess_g = find_best_col(gastos_raw, POSSIBLE_CC)
cc_guess_b = find_best_col(budget_raw, POSSIBLE_CC) if budget_raw is not None else None
budget_fy_guess = find_best_col(budget_raw, POSSIBLE_BUDGET_FY) if budget_raw is not None else None

classif_col = st.sidebar.selectbox("Columna de naturaleza/clasificación", all_cols, index=all_cols.index(classif_guess) if classif_guess in all_cols else 0)
cc_col_gastos = st.sidebar.selectbox("Llave en gastos (CC/Resp/Centro costo)", all_cols, index=all_cols.index(cc_guess_g) if cc_guess_g in all_cols else 0)
if budget_raw is not None:
    budget_cols_options = [None] + list(budget_raw.columns)
    cc_col_budget = st.sidebar.selectbox("Llave en Budget", budget_cols_options, index=budget_cols_options.index(cc_guess_b) if cc_guess_b in budget_cols_options else 0)
    budget_fy_col = st.sidebar.selectbox("Columna Budget FY total (opcional)", budget_cols_options, index=budget_cols_options.index(budget_fy_guess) if budget_fy_guess in budget_cols_options else 0)
else:
    cc_col_budget, budget_fy_col = None, None

candidate_dims = [c for c in gastos_raw.columns if c not in gastos_month_cols.values()]
default_dims = [c for c in candidate_dims if norm_txt(c) in ["vp", "gerencia", "desc resp", "desc proc", "classif", "naturaleza"]]
extra_dims = st.sidebar.multiselect("Filtros/dimensiones para dashboard", candidate_dims, default=default_dims[:6])

st.sidebar.subheader("4) Parámetros del modelo")
sensitivity_mult = st.sidebar.slider("Sensibilidad a ejecución YTD", 0.2, 2.0, 1.0, 0.05)
momentum_mult = st.sidebar.slider("Peso de tendencia reciente", 0.0, 2.0, 1.0, 0.05)
scenario_mult = st.sidebar.slider("Intensidad curva no lineal", 0.0, 2.0, 1.0, 0.05)
steepness = st.sidebar.slider("Forma de curva logística", 0.5, 3.0, 1.35, 0.05)

st.sidebar.subheader("5) Simulación Estratégica (5 Años)")
st.sidebar.markdown("Define el escenario macroeconómico y operativo:")
cagr_inf = st.sidebar.slider("Inflación Anual Estimada (%)", 0.0, 10.0, 3.0, 0.5)
cagr_ops = st.sidebar.slider("Crecimiento Operacional Anual (%)", -5.0, 15.0, 2.0, 0.5)

try:
    prepared = prepare_data(
        gastos_raw, budget_raw, gastos_month_cols, budget_month_cols,
        actual_months, forecast_months, classif_col, cc_col_gastos, cc_col_budget,
        budget_fy_col, extra_dims,
    )
    model = calculate_forecast(prepared, actual_months, forecast_months, sensitivity_mult, momentum_mult, scenario_mult, steepness)
except Exception as e:
    st.error(f"Error al preparar/calcular el forecast: {e}")
    st.stop()

# Filtros dinámicos
filtered = model.copy()
with st.sidebar.expander("6) Filtros", expanded=True):
    if "Naturaleza" in filtered.columns:
        selected_nat = st.multiselect("Naturaleza", sorted(filtered["Naturaleza"].dropna().astype(str).unique()), default=[])
        if selected_nat:
            filtered = filtered[filtered["Naturaleza"].astype(str).isin(selected_nat)]
    for dim in extra_dims[:8]:
        if dim in filtered.columns:
            vals = sorted(filtered[dim].dropna().astype(str).unique())
            if 1 < len(vals) <= 200:
                sel = st.multiselect(str(dim), vals, default=[])
                if sel:
                    filtered = filtered[filtered[dim].astype(str).isin(sel)]

# ─────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
actual_total = filtered["Actual_YTD"].sum()
budget_total = filtered["Budget_FY_Model"].sum()
forecast_total = filtered["Forecast_FY_Modelo"].sum()
var_total = forecast_total - budget_total
var_pct = safe_div(var_total, budget_total, 0)
kpi1.metric("Actual YTD", money(actual_total))
kpi2.metric("Budget FY", money(budget_total))
kpi3.metric("Forecast FY Modelo", money(forecast_total))
kpi4.metric("Var vs Budget", money(var_total), f"{var_pct:.1%}")

st.divider()

# Agrupación para gráficos
chart_dim_options = ["Naturaleza", "Contexto_Mina"] + [c for c in extra_dims if c in filtered.columns]
chart_dim = st.selectbox("Agrupar dashboard por", chart_dim_options, index=0)
summary = aggregate_for_charts(filtered, chart_dim)

# ═══════════════════════════════════════════════════════════════
# PESTAÑAS PRINCIPALES
# ═══════════════════════════════════════════════════════════════
tab_actual, tab_5y = st.tabs(["📊 Forecast Año Actual (5+7)", "🚀 Simulación LRP a 5 Años"])

# ─────────────────────────────────────────────
# PESTAÑA 1: FORECAST AÑO ACTUAL
# ─────────────────────────────────────────────
with tab_actual:

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(summary.head(15), x=chart_dim, y=["Budget_FY_Model", "Forecast_FY_Modelo"], barmode="group", title="Forecast FY vs Budget FY")
        fig.update_layout(xaxis_title="Categoría", yaxis_title="Monto")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig2 = px.bar(summary.sort_values("Var_vs_Budget", ascending=False).head(15), x=chart_dim, y="Var_vs_Budget", title="Principales desviaciones vs Budget")
        fig2.update_layout(xaxis_title="Categoría", yaxis_title="Varianza")
        st.plotly_chart(fig2, use_container_width=True)

    # Curva mensual total
    monthly_rows = []
    for m in MONTH_ORDER:
        if m in actual_months:
            monthly_rows.append({"Mes": m, "Tipo": "Actual", "Monto": filtered.get(f"Actual_{m}", pd.Series(0, index=filtered.index)).sum()})
        elif m in forecast_months:
            monthly_rows.append({"Mes": m, "Tipo": "Forecast Modelo", "Monto": filtered.get(f"Forecast_{m}", pd.Series(0, index=filtered.index)).sum()})
        monthly_rows.append({"Mes": m, "Tipo": "Budget", "Monto": filtered.get(f"Budget_{m}", pd.Series(0, index=filtered.index)).sum()})
    monthly_df = pd.DataFrame(monthly_rows)
    fig3 = px.line(monthly_df, x="Mes", y="Monto", color="Tipo", markers=True, title="Serie mensual: Actual + Forecast no lineal vs Budget")
    st.plotly_chart(fig3, use_container_width=True)

    st.subheader("📌 Hallazgos automáticos")
    if len(summary) > 0:
        top_spend = summary.iloc[0]
        top_over = summary.sort_values("Var_vs_Budget", ascending=False).iloc[0]
        top_under = summary.sort_values("Var_vs_Budget", ascending=True).iloc[0]
        st.markdown(f"""
        - La mayor concentración de gasto proyectado está en **{top_spend[chart_dim]}**, con un forecast de **{money(top_spend['Forecast_FY_Modelo'])}**.
        - La mayor desviación positiva contra presupuesto está en **{top_over[chart_dim]}**, con **{money(top_over['Var_vs_Budget'])}** sobre Budget.
        - La mayor subejecución proyectada está en **{top_under[chart_dim]}**, con **{money(top_under['Var_vs_Budget'])}** respecto al Budget.
        - El forecast total del filtro seleccionado queda en **{money(forecast_total)}**, equivalente a una desviación de **{var_pct:.1%}** contra el presupuesto anual.
        """)

    # Gráfico Waterfall
    st.subheader("🌊 Gráfico cascada: explicación de la variación presupuestaria")
    waterfall_df = summary.copy()
    top_over_wf = waterfall_df[waterfall_df["Var_vs_Budget"] > 0].sort_values("Var_vs_Budget", ascending=False).head(3)
    top_under_wf = waterfall_df[waterfall_df["Var_vs_Budget"] < 0].sort_values("Var_vs_Budget", ascending=True).head(3)
    waterfall_items = pd.concat([top_over_wf, top_under_wf])

    inicio = budget_total
    final = forecast_total
    labels = ["Budget FY"] + waterfall_items[chart_dim].astype(str).tolist() + ["Forecast FY"]
    values = [inicio] + waterfall_items["Var_vs_Budget"].tolist() + [final]
    measure = ["absolute"] + ["relative"] * len(waterfall_items) + ["total"]

    fig_waterfall = go.Figure(go.Waterfall(
        name="Variación", orientation="v", measure=measure, x=labels, y=values,
        text=[money(inicio)] + [money(v) for v in waterfall_items["Var_vs_Budget"]] + [money(final)],
        textposition="outside",
        connector={"line": {"color": "gray"}},
        increasing={"marker": {"color": "#4C78A8"}},
        decreasing={"marker": {"color": "#E45756"}},
        totals={"marker": {"color": "#72B7B2"}},
    ))
    fig_waterfall.update_layout(
        title="Explicación de la variación entre Budget FY y Forecast FY",
        yaxis_title="Monto", xaxis_title="Categoría", showlegend=False, height=520,
    )
    st.plotly_chart(fig_waterfall, use_container_width=True)

    if len(waterfall_items) > 0:
        top_pos = waterfall_items[waterfall_items["Var_vs_Budget"] > 0].sort_values("Var_vs_Budget", ascending=False)
        top_neg = waterfall_items[waterfall_items["Var_vs_Budget"] < 0].sort_values("Var_vs_Budget", ascending=True)
        texto_pos = f"- La mayor desviación positiva corresponde a **{top_pos.iloc[0][chart_dim]}**, con un aumento de **{money(top_pos.iloc[0]['Var_vs_Budget'])}** respecto al presupuesto.\n" if len(top_pos) > 0 else "- No se observan categorías con desviación positiva relevante.\n"
        texto_neg = f"- La mayor desviación negativa corresponde a **{top_neg.iloc[0][chart_dim]}**, con una reducción de **{money(top_neg.iloc[0]['Var_vs_Budget'])}** respecto al presupuesto.\n" if len(top_neg) > 0 else "- No se observan categorías con desviación negativa relevante.\n"
        conclusion_var = "El modelo proyecta una **subejecución presupuestaria**." if var_total < 0 else ("El modelo proyecta una **sobreejecución presupuestaria**." if var_total > 0 else "El modelo proyecta un cierre alineado con el presupuesto.")
        st.markdown(f"""
### Interpretación automática
- Budget FY: **{money(budget_total)}** | Forecast FY Modelo: **{money(forecast_total)}** | Desviación: **{money(var_total)} ({var_pct:.1%})**
{texto_pos}{texto_neg}{conclusion_var}
        """)

    st.subheader("✅ Propuesta formal de mejora")
    st.markdown("Se recomienda implementar un sistema de control presupuestario con alertas tempranas por naturaleza de gasto y centro de costo. Las partidas con desviaciones positivas deben revisarse mediante acciones de control operacional, renegociación contractual y revisión de consumos críticos.")

    st.subheader("📄 Resultado detallado")
    cols_to_show = [c for c in extra_dims if c in filtered.columns] + ["Naturaleza", "Contexto_Mina", "Actual_YTD", "Budget_YTD", "Budget_Remaining", "Forecast_Remaining", "Budget_FY_Model", "Forecast_FY_Modelo", "Var_vs_Budget", "Var_vs_Budget_%", "Recomendacion", "Justificacion_Mina"]
    cols_to_show = list(dict.fromkeys([c for c in cols_to_show if c in filtered.columns]))
    st.dataframe(filtered[cols_to_show].sort_values("Var_vs_Budget", ascending=False), use_container_width=True, height=420)

    excel_bytes = to_excel_bytes(filtered[cols_to_show], summary)
    st.download_button(
        "⬇️ Descargar Forecast en Excel",
        data=excel_bytes,
        file_name="forecast_5mas7_no_lineal_resultados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─────────────────────────────────────────────
# PESTAÑA 2: SIMULACIÓN LRP 5 AÑOS
# ─────────────────────────────────────────────
with tab_5y:
    año_base = datetime.datetime.now().year

    st.subheader("Simulación a 5 Años basada en Sensibilidad Operativa")
    st.caption(f"Inflación: {cagr_inf}% | Crecimiento operacional: {cagr_ops}% | Base: Forecast FY Modelo")

    if len(filtered) > 0:
        df_5y = simular_5_anos(filtered, cagr_inf, cagr_ops)

        cols_agrup = ["Contexto_Mina"]
        if "Naturaleza" in df_5y.columns:
            cols_agrup = ["Contexto_Mina", "Naturaleza"]

        año_cols = ["Año_0_Base", "Año_1", "Año_2", "Año_3", "Año_4", "Año_5"]
        resumen_5y_raw = df_5y.groupby(cols_agrup)[año_cols].sum().reset_index()

        # Estacionalidad para desglosar Año_1 en 12 meses
        estacionalidad = calcular_estacionalidad_mensual(filtered, actual_months, forecast_months)
        año1_label = año_base + 1
        meses_año1 = [f"{m}-{str(año1_label)[2:]}" for m in MONTH_ORDER]

        for m, mes_col in zip(MONTH_ORDER, meses_año1):
            resumen_5y_raw[mes_col] = resumen_5y_raw["Año_1"] * estacionalidad[m]

        # Renombrar años acumulados (Año_0_Base = FY26, Año_2..5 = FY28..31)
        rename_map = {
            "Año_0_Base": f"FY{str(año_base)[2:]}",
            "Año_2": f"FY{str(año_base + 2)[2:]}",
            "Año_3": f"FY{str(año_base + 3)[2:]}",
            "Año_4": f"FY{str(año_base + 4)[2:]}",
            "Año_5": f"FY{str(año_base + 5)[2:]}",
        }
        resumen_5y = resumen_5y_raw.rename(columns=rename_map).drop(columns=["Año_1"])

        fy_acumulados = [f"FY{str(año_base + i)[2:]}" for i in [0, 2, 3, 4, 5]]
        cols_orden = cols_agrup + [fy_acumulados[0]] + meses_año1 + fy_acumulados[1:]
        resumen_5y = resumen_5y[[c for c in cols_orden if c in resumen_5y.columns]]

        # Fila TOTAL
        cols_money = [c for c in resumen_5y.columns if c not in cols_agrup]
        total_row = {}
        for i, col in enumerate(resumen_5y.columns):
            if col in cols_money:
                total_row[col] = resumen_5y[col].sum()
            elif i == 0:
                total_row[col] = "TOTAL"
            else:
                total_row[col] = ""
        resumen_5y_con_total = pd.concat([resumen_5y, pd.DataFrame([total_row])], ignore_index=True)

        def fmt_mm(val):
            try:
                return f"US$ {float(val)/1_000_000:,.2f} MM"
            except Exception:
                return val

        resumen_display = resumen_5y_con_total.copy()
        for col in cols_money:
            resumen_display[col] = resumen_display[col].apply(fmt_mm)

        st.markdown(f"**Proyección LRP — {año1_label} mensual + acumulados anuales (millones USD)**")
        st.dataframe(resumen_display, use_container_width=True, hide_index=True)

        # Gráfico de áreas (sin FY26 base, arranca en FY27)
        año_cols_sin_base = ["Año_1", "Año_2", "Año_3", "Año_4", "Año_5"]
        año_cols_grafico  = [f"FY{str(año_base + i)[2:]}" for i in range(1, 6)]
        resumen_grafico = df_5y.groupby(cols_agrup)[año_cols_sin_base].sum().reset_index()
        resumen_grafico = resumen_grafico.rename(columns=dict(zip(año_cols_sin_base, año_cols_grafico)))
        resumen_largo = resumen_grafico.melt(
            id_vars=cols_agrup, var_name="Año", value_name="Presupuesto Proyectado"
        )

        fig_5y = px.area(
            resumen_largo,
            x="Año",
            y="Presupuesto Proyectado",
            color="Contexto_Mina",
            title=f"Evolución Estructural del Presupuesto FY{str(año_base+1)[2:]}–FY{str(año_base+5)[2:]}",
            markers=True,
        )
        st.plotly_chart(fig_5y, use_container_width=True)

        # Botón descarga reporte completo
        excel_completo = to_excel_completo(
            filtered, summary, resumen_5y, resumen_largo, actual_months, forecast_months
        )
        st.download_button(
            label="⬇️ Descargar Reporte Completo: Forecast 5+7 + Simulación LRP (Excel)",
            data=excel_completo,
            file_name="reporte_forecast_lrp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    else:
        st.warning("⚠️ No hay datos para simular. Ajusta los filtros en el menú lateral.")

st.caption("Modelo diseñado para defensa académica: combina ejecución acumulada, tendencia reciente, curva no lineal y contexto minero.")
