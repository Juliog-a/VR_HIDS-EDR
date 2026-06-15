#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
TFM - Procesado automático de benchmarks de Velociraptor (logman CSV -> Excel)
Modelo: 4 escenarios

Escenarios esperados:
    1. BASE_SIN_VR
    2. BASE_CON_VR
    3. TEC_SIN_VR
    4. TEC_CON_VR

Nombres recomendados:
    BENCH_BASE_SIN_VR_R1.csv
    BENCH_BASE_CON_VR_R1.csv
    BENCH_TEC-001_SIN_VR_R1.csv
    BENCH_TEC-001_CON_VR_R1.csv

También acepta variantes razonables:
    TEC001 / TEC-001
    VR_ON / CON_VR / ON
    VR_OFF / SIN_VR / OFF
    BASELINE / BASE
    MEDIDA / MEASURE  (compatibilidad con formato antiguo)

Dependencias:
    pip install pandas xlsxwriter

Uso:
    python benchmark_vr_to_excel_4escenarios.py \
        --input-dir "C:\\Users\\seguridad\\Desktop\\TFM\\TFM_Benchmark\\Logs" \
        --output-xlsx "C:\\Users\\seguridad\\Desktop\\TFM\\TFM_Benchmark\\TFM_Benchmark_Resumen.xlsx"
"""

from __future__ import annotations

import argparse
import math
import re
import unicodedata
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd


# ============================================================
# Configuración por defecto
# ============================================================

DEFAULT_INPUT_DIR = Path(r"C:\Users\seguridad\Desktop\TFM\TFM_Benchmark\Logs")
DEFAULT_OUTPUT_XLSX = DEFAULT_INPUT_DIR / "TFM_Benchmark_Resumen.xlsx"
DEFAULT_SAMPLE_INTERVAL_S = 1
DEFAULT_TOOL_NAME = "Performance Monitor (logman)"
DEFAULT_PROCESS_NAME = "velociraptor"
DEFAULT_MIN_DURATION_S = 30

SCENARIOS = ["BASE_SIN_VR", "BASE_CON_VR", "TEC_SIN_VR", "TEC_CON_VR"]

CORE_METRICS = [
    "CPU_Total",
    "RAM_Available_MB",
    "IO_Reads_sec",
    "IO_Writes_sec",
    "VR_CPU",
    "VR_RAM_MB",
    "VR_IO_Read_Bps",
    "VR_IO_Write_Bps",
]

STAT_SUFFIXES = ["Mean", "Max", "Min", "Median", "STD", "Samples"]


# ============================================================
# Utilidades generales
# ============================================================

def normalize_text(value: object) -> str:
    """Normaliza texto para comparar cabeceras y nombres de archivo."""
    txt = str(value).strip().lower()
    txt = txt.replace('"', "")
    txt = txt.replace("\\\\", "\\")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"\s+", " ", txt)
    return txt


def normalize_token(value: object) -> str:
    """Normaliza tokens de nombre de archivo a mayúsculas con separador _."""
    txt = str(value).strip().upper()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"[^A-Z0-9]+", "_", txt)
    txt = re.sub(r"_+", "_", txt).strip("_")
    return txt


def safe_float(value: object) -> Optional[float]:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        return float(value)
    except Exception:
        return None


def clean_number_series(series: pd.Series) -> pd.Series:
    """
    Convierte una serie a numérica.
    Soporta salida logman con punto decimal y, si falla, coma decimal.
    """
    numeric = pd.to_numeric(series, errors="coerce")
    if numeric.notna().sum() > 0:
        return numeric

    as_text = series.astype(str).str.strip()
    as_text = as_text.str.replace("\u00a0", "", regex=False)
    as_text = as_text.str.replace(".", "", regex=False)
    as_text = as_text.str.replace(",", ".", regex=False)
    return pd.to_numeric(as_text, errors="coerce")


def numeric_series(df: pd.DataFrame, col: str) -> pd.Series:
    return clean_number_series(df[col])


def bytes_to_mb_series(series: pd.Series) -> pd.Series:
    return clean_number_series(series) / (1024 * 1024)


def round_or_none(value: object, ndigits: int = 3) -> Optional[float]:
    v = safe_float(value)
    if v is None:
        return None
    if math.isnan(v) or math.isinf(v):
        return None
    return round(v, ndigits)


def delta(measured: object, base: object) -> Optional[float]:
    """Delta bruto: medida - base."""
    m = safe_float(measured)
    b = safe_float(base)
    if m is None or b is None:
        return None
    return m - b


def ram_impact(base_available: object, measured_available: object) -> Optional[float]:
    """
    RAM disponible -> impacto interpretativo.
    Positivo = más consumo de RAM.
    Fórmula: base disponible - medida disponible.
    """
    b = safe_float(base_available)
    m = safe_float(measured_available)
    if b is None or m is None:
        return None
    return b - m


def clamp_positive(value: object) -> Optional[float]:
    v = safe_float(value)
    if v is None:
        return None
    return max(0.0, v)


def aggregate_mean_std(series: pd.Series) -> Tuple[Optional[float], Optional[float], int]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return None, None, 0
    std = float(s.std(ddof=1)) if len(s) > 1 else 0.0
    return float(s.mean()), std, int(s.count())


def stats_from_series(series: pd.Series) -> Dict[str, Optional[float]]:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if s.empty:
        return {"Mean": None, "Max": None, "Min": None, "Median": None, "STD": None, "Samples": 0}
    return {
        "Mean": float(s.mean()),
        "Max": float(s.max()),
        "Min": float(s.min()),
        "Median": float(s.median()),
        "STD": float(s.std(ddof=1)) if len(s) > 1 else 0.0,
        "Samples": int(s.count()),
    }


def parse_timestamp_series(series: pd.Series) -> pd.Series:
    """Parsea la primera columna de CSV logman."""
    ts = pd.to_datetime(series, errors="coerce", dayfirst=False)
    if ts.notna().sum() == 0:
        ts = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return ts


# ============================================================
# Detección de columnas logman / PerfMon
# ============================================================

def header_matches(header: str, alias_groups: Sequence[Sequence[str] | str]) -> bool:
    h = normalize_text(header)
    for group in alias_groups:
        aliases = [group] if isinstance(group, str) else list(group)
        aliases_norm = [normalize_text(alias) for alias in aliases]
        if not any(alias in h for alias in aliases_norm):
            return False
    return True


def matching_columns(columns: Sequence[str], alias_groups: Sequence[Sequence[str] | str]) -> List[str]:
    return [col for col in columns if header_matches(col, alias_groups)]


def prefer_total_column(cols: List[str]) -> List[str]:
    if not cols:
        return []
    total = [c for c in cols if "(_total)" in normalize_text(c) or "(_total" in normalize_text(c)]
    return total if total else cols


def first_matching_column(columns: Sequence[str], alias_groups: Sequence[Sequence[str] | str]) -> Optional[str]:
    matches = prefer_total_column(matching_columns(columns, alias_groups))
    return matches[0] if matches else None


def process_matching_columns(columns: Sequence[str], metric_aliases: Sequence[str] | str, process_name: str) -> List[str]:
    r"""
    Selecciona columnas de proceso tipo:
      \\HOST\Process(velociraptor)\% Processor Time
      \\HOST\Proceso(velociraptor)\Espacio de trabajo - Privado
    """
    aliases = [metric_aliases] if isinstance(metric_aliases, str) else list(metric_aliases)
    aliases_norm = [normalize_text(x) for x in aliases]
    proc = normalize_text(process_name)

    out: List[str] = []
    for col in columns:
        h = normalize_text(col)
        if "\\process(" not in h and "\\proceso(" not in h:
            continue
        if proc not in h:
            continue
        if not any(alias in h for alias in aliases_norm):
            continue
        out.append(col)
    return out


def aggregate_columns_sum(df: pd.DataFrame, cols: Sequence[str]) -> pd.Series:
    if not cols:
        return pd.Series(dtype="float64")
    temp = pd.DataFrame({col: clean_number_series(df[col]) for col in cols})
    return temp.sum(axis=1, min_count=1)


# ============================================================
# Inferencia de metadatos desde nombre de archivo
# ============================================================

def normalize_tech(raw: str) -> str:
    m = re.search(r"TEC[-_ ]?(\d{3})", raw.upper())
    if not m:
        return raw.upper()
    return f"TEC-{m.group(1)}"


def infer_metadata_from_filename(file_name: str) -> Dict[str, object]:
    """
    Devuelve metadatos normalizados desde el nombre del CSV.

    Formato recomendado:
        BENCH_BASE_SIN_VR_R1.csv
        BENCH_BASE_CON_VR_R1.csv
        BENCH_TEC-001_SIN_VR_R1.csv
        BENCH_TEC-001_CON_VR_R1.csv

    Compatibilidad antigua:
        BENCH_BASE_R1.csv            -> BASE_SIN_VR (asunción)
        BENCH_TEC-001_MEDIDA_R1.csv  -> TEC_CON_VR  (asunción)
    """
    stem_raw = Path(file_name).stem.upper()
    stem = normalize_token(stem_raw)
    warnings: List[str] = []

    tech: Optional[str] = None
    m_tech = re.search(r"TEC[_\- ]?(\d{3})", stem_raw)
    if m_tech:
        tech = f"TEC-{m_tech.group(1)}"

    repetition: Optional[int] = None
    m_rep = re.search(r"(?:^|[_\-\s])R(\d+)(?:[_\-\s]|$)", stem_raw)
    if m_rep:
        repetition = int(m_rep.group(1))

    has_base = bool(re.search(r"(?:^|_)BASE(?:LINE)?(?:_|$)", stem)) or "BASELINE" in stem
    has_measure_old = bool(re.search(r"(?:^|_)(MEDIDA|MEASURE|LOAD|CARGA)(?:_|$)", stem))

    tipo_carga: Optional[str] = None
    if tech:
        tipo_carga = "TEC"
    elif has_base:
        tipo_carga = "BASE"
    elif has_measure_old:
        tipo_carga = "TEC"
        warnings.append("Formato antiguo MEDIDA/MEASURE sin técnica detectada.")

    vr_status: Optional[str] = None

    # Casos explícitos recomendados.
    if re.search(r"(?:^|_)CON_VR(?:_|$)", stem) or re.search(r"(?:^|_)VR_ON(?:_|$)", stem):
        vr_status = "CON_VR"
    elif re.search(r"(?:^|_)SIN_VR(?:_|$)", stem) or re.search(r"(?:^|_)VR_OFF(?:_|$)", stem):
        vr_status = "SIN_VR"

    # Variantes compactas.
    if vr_status is None:
        if re.search(r"(?:^|_)(VRON|ONVR|WITHVR|WITH_VR)(?:_|$)", stem):
            vr_status = "CON_VR"
        elif re.search(r"(?:^|_)(VROFF|OFFVR|WITHOUTVR|WITHOUT_VR)(?:_|$)", stem):
            vr_status = "SIN_VR"

    # Compatibilidad con formato antiguo.
    if vr_status is None and tipo_carga == "BASE" and has_base:
        vr_status = "SIN_VR"
        warnings.append("Formato antiguo BASE sin estado VR: se asume BASE_SIN_VR.")
    elif vr_status is None and tipo_carga == "TEC" and has_measure_old:
        vr_status = "CON_VR"
        warnings.append("Formato antiguo MEDIDA/MEASURE sin estado VR: se asume TEC_CON_VR.")

    escenario_normalizado: Optional[str] = None
    if tipo_carga and vr_status:
        escenario_normalizado = f"{tipo_carga}_{vr_status}"

    bench_id = None
    if escenario_normalizado:
        bench_id = f"BENCH_{tech}_{vr_status}" if tech else f"BENCH_{escenario_normalizado}"

    metadata_completa = bool(
        escenario_normalizado
        and repetition is not None
        and (tipo_carga == "BASE" or (tipo_carga == "TEC" and tech))
    )

    return {
        "ID_Benchmark_Filename": bench_id,
        "ID_Tecnica_Interna": tech,
        "Tipo_Carga": tipo_carga,
        "Estado_VR": vr_status,
        "Escenario_Normalizado": escenario_normalizado,
        "Repeticion": repetition,
        "Metadata_Completa": "Sí" if metadata_completa else "No",
        "Advertencias_Metadata": " | ".join(warnings),
    }


# ============================================================
# Lectura y extracción de métricas por CSV
# ============================================================

def read_logman_csv(csv_path: Path) -> pd.DataFrame:
    errors = []
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            df = pd.read_csv(csv_path, encoding=enc, engine="python", sep=None)
            if df.shape[1] == 1:
                # Fallback habitual si el detector falla.
                df = pd.read_csv(csv_path, encoding=enc, engine="python", sep=",")
            return df
        except Exception as exc:
            errors.append(f"{enc}: {exc}")
    raise RuntimeError("No se pudo leer el CSV. Encodings probados: " + " || ".join(errors))


def extract_metrics_from_file(csv_path: Path, process_name: str) -> Dict[str, object]:
    df = read_logman_csv(csv_path)
    cols = list(df.columns)
    meta = infer_metadata_from_filename(csv_path.name)

    row: Dict[str, object] = {
        "Archivo": csv_path.name,
        "Ruta": str(csv_path),
        "Filas_CSV": len(df),
        "Columnas_CSV": len(df.columns),
        **meta,
    }

    # Tiempo / duración.
    first_col = cols[0] if cols else None
    if first_col:
        ts = parse_timestamp_series(df[first_col])
        valid_ts = ts.dropna()
        row["Muestras_Tiempo_Validas"] = int(valid_ts.count())
        if not valid_ts.empty:
            t0 = valid_ts.min()
            t1 = valid_ts.max()
            row["Timestamp_Inicio"] = t0.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            row["Timestamp_Fin"] = t1.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            row["Duracion_Real_s"] = round((t1 - t0).total_seconds(), 3)
        else:
            row["Timestamp_Inicio"] = None
            row["Timestamp_Fin"] = None
            row["Duracion_Real_s"] = None
    else:
        row["Muestras_Tiempo_Validas"] = 0
        row["Timestamp_Inicio"] = None
        row["Timestamp_Fin"] = None
        row["Duracion_Real_s"] = None

    # Métricas globales.
    cpu_col = first_matching_column(cols, [
        ["processor(_total)", "procesador(_total)"],
        ["% processor time", "% de tiempo de procesador"],
    ])
    ram_col = first_matching_column(cols, [
        ["memory", "memoria"],
        ["available mbytes", "mbytes disponibles", "mb disponibles"],
    ])
    read_col = first_matching_column(cols, [
        ["physicaldisk(_total)", "disco fisico(_total)", "discofisico(_total)"],
        ["disk reads/sec", "lecturas de disco/s", "lecturas disco/s"],
    ])
    write_col = first_matching_column(cols, [
        ["physicaldisk(_total)", "disco fisico(_total)", "discofisico(_total)"],
        ["disk writes/sec", "escrituras en disco/s", "escrituras disco/s"],
    ])

    row["Cols_CPU_Total"] = cpu_col or ""
    row["Cols_RAM_Available"] = ram_col or ""
    row["Cols_IO_Reads"] = read_col or ""
    row["Cols_IO_Writes"] = write_col or ""

    metric_series: Dict[str, pd.Series] = {
        "CPU_Total": numeric_series(df, cpu_col) if cpu_col else pd.Series(dtype="float64"),
        "RAM_Available_MB": numeric_series(df, ram_col) if ram_col else pd.Series(dtype="float64"),
        "IO_Reads_sec": numeric_series(df, read_col) if read_col else pd.Series(dtype="float64"),
        "IO_Writes_sec": numeric_series(df, write_col) if write_col else pd.Series(dtype="float64"),
    }

    # Métricas proceso Velociraptor.
    vr_cpu_cols = process_matching_columns(cols, ["% processor time", "% de tiempo de procesador"], process_name)
    vr_ram_cols = process_matching_columns(cols, [
        "working set - private",
        "private working set",
        "espacio de trabajo - privado",
        "conjunto de trabajo privado",
    ], process_name)
    vr_read_cols = process_matching_columns(cols, [
        "io read bytes/sec",
        "bytes de lectura de es/s",
        "bytes lectura es/s",
    ], process_name)
    vr_write_cols = process_matching_columns(cols, [
        "io write bytes/sec",
        "bytes de escritura de es/s",
        "bytes escritura es/s",
    ], process_name)

    row["Cols_VR_CPU"] = "; ".join(vr_cpu_cols)
    row["Cols_VR_RAM"] = "; ".join(vr_ram_cols)
    row["Cols_VR_IO_Read"] = "; ".join(vr_read_cols)
    row["Cols_VR_IO_Write"] = "; ".join(vr_write_cols)

    metric_series.update({
        "VR_CPU": aggregate_columns_sum(df, vr_cpu_cols),
        "VR_RAM_MB": bytes_to_mb_series(aggregate_columns_sum(df, vr_ram_cols)),
        "VR_IO_Read_Bps": aggregate_columns_sum(df, vr_read_cols),
        "VR_IO_Write_Bps": aggregate_columns_sum(df, vr_write_cols),
    })

    for metric, series in metric_series.items():
        stats = stats_from_series(series)
        for suffix in STAT_SUFFIXES:
            row[f"{metric}_{suffix}"] = stats[suffix]

    return row


# ============================================================
# Control de calidad
# ============================================================

def build_quality_control(files_df: pd.DataFrame,
                          errors_df: pd.DataFrame,
                          min_duration_s: int) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []

    for _, r in files_df.iterrows():
        warnings: List[str] = []
        escenario = r.get("Escenario_Normalizado")
        tipo = r.get("Tipo_Carga")
        tech = r.get("ID_Tecnica_Interna")
        rep = r.get("Repeticion")
        estado_vr = r.get("Estado_VR")

        if not escenario:
            warnings.append("falta escenario normalizado")
        if tipo == "TEC" and not tech:
            warnings.append("falta técnica")
        if pd.isna(rep):
            warnings.append("falta repetición")
        if r.get("Metadata_Completa") != "Sí":
            warnings.append("metadata incompleta")

        duration = safe_float(r.get("Duracion_Real_s"))
        if duration is None:
            warnings.append("duración no calculada")
        elif duration < min_duration_s:
            warnings.append(f"duración demasiado corta (<{min_duration_s}s)")

        samples = safe_float(r.get("Filas_CSV"))
        if samples is None or samples < 2:
            warnings.append("número de muestras insuficiente")

        missing_cols = []
        if not r.get("Cols_CPU_Total"):
            missing_cols.append("CPU total")
        if not r.get("Cols_RAM_Available"):
            missing_cols.append("RAM disponible")
        if not r.get("Cols_IO_Reads"):
            missing_cols.append("Disk Reads/sec")
        if not r.get("Cols_IO_Writes"):
            missing_cols.append("Disk Writes/sec")
        if missing_cols:
            warnings.append("columnas no encontradas: " + ", ".join(missing_cols))

        vr_samples = [
            safe_float(r.get("VR_CPU_Samples")) or 0,
            safe_float(r.get("VR_RAM_MB_Samples")) or 0,
            safe_float(r.get("VR_IO_Read_Bps_Samples")) or 0,
            safe_float(r.get("VR_IO_Write_Bps_Samples")) or 0,
        ]
        if estado_vr == "CON_VR" and max(vr_samples) == 0:
            warnings.append("sin métricas del proceso Velociraptor en escenario CON_VR")

        # En escenarios SIN_VR, las métricas específicas de Velociraptor no se evalúan.
        # La comparación SIN_VR se basa en métricas globales del sistema: CPU, RAM e I/O.
        if estado_vr == "CON_VR":
            metricas_vr_estado = "Sí" if max(vr_samples) > 0 else "No"
        elif estado_vr == "SIN_VR":
            metricas_vr_estado = "No aplica"
        else:
            metricas_vr_estado = "No evaluado"

        meta_warn = str(r.get("Advertencias_Metadata") or "").strip()
        if meta_warn:
            warnings.append(meta_warn)

        rows.append({
            "Archivo": r.get("Archivo"),
            "Procesado": "Sí",
            "Error_Procesado": "No",
            "Metadata_Completa": r.get("Metadata_Completa"),
            "Escenario_Normalizado": escenario,
            "ID_Tecnica_Interna": tech,
            "Estado_VR": estado_vr,
            "Repeticion": rep,
            "Duracion_Real_s": r.get("Duracion_Real_s"),
            "Numero_Muestras": r.get("Filas_CSV"),
            "Columnas_CSV": r.get("Columnas_CSV"),
            "Columnas_CPU_RAM_IO_OK": "Sí" if not missing_cols else "No",
            "Metricas_VR_Encontradas": metricas_vr_estado,
            "Advertencias": " | ".join(warnings) if warnings else "No",
        })

    if errors_df is not None and not errors_df.empty:
        for _, e in errors_df.iterrows():
            rows.append({
                "Archivo": e.get("Archivo"),
                "Procesado": "No",
                "Error_Procesado": "Sí",
                "Metadata_Completa": "No",
                "Escenario_Normalizado": None,
                "ID_Tecnica_Interna": None,
                "Estado_VR": None,
                "Repeticion": None,
                "Duracion_Real_s": None,
                "Numero_Muestras": None,
                "Columnas_CSV": None,
                "Columnas_CPU_RAM_IO_OK": "No",
                "Metricas_VR_Encontradas": "No",
                "Advertencias": "CSV con error de lectura/procesado",
            })

    return pd.DataFrame(rows).sort_values(by=["Archivo"]).reset_index(drop=True)


# ============================================================
# Agregaciones base y escenarios
# ============================================================

def aggregate_group(group: pd.DataFrame, label_cols: Dict[str, object]) -> Dict[str, object]:
    out: Dict[str, object] = dict(label_cols)
    out["N_Repeticiones"] = int(len(group))
    out["Archivos"] = "; ".join(group["Archivo"].dropna().astype(str).tolist())

    dur_mean, dur_std, _ = aggregate_mean_std(group["Duracion_Real_s"])
    out["Duracion_Media_s"] = dur_mean
    out["Duracion_STD_s"] = dur_std

    for metric in CORE_METRICS:
        source_col = f"{metric}_Mean"
        if source_col in group.columns:
            mean, std, n = aggregate_mean_std(group[source_col])
        else:
            mean, std, n = None, None, 0
        out[f"{metric}_Media"] = mean
        out[f"{metric}_STD_Entre_Reps"] = std
        out[f"{metric}_N"] = n

    return out


def build_baselines(files_df: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    if files_df.empty:
        return pd.DataFrame()

    for scenario in ["BASE_SIN_VR", "BASE_CON_VR"]:
        g = files_df[files_df["Escenario_Normalizado"] == scenario].copy()
        if g.empty:
            rows.append({
                "Escenario_Normalizado": scenario,
                "N_Repeticiones": 0,
                "Estado": "Sin datos",
            })
        else:
            row = aggregate_group(g, {"Escenario_Normalizado": scenario, "Estado": "OK"})
            rows.append(row)

    return pd.DataFrame(rows)


def build_scenario_aggregates(files_df: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    if files_df.empty:
        return pd.DataFrame()

    for scenario in SCENARIOS:
        g_s = files_df[files_df["Escenario_Normalizado"] == scenario].copy()
        if g_s.empty:
            continue

        if scenario.startswith("BASE"):
            rows.append(aggregate_group(g_s, {
                "Escenario_Normalizado": scenario,
                "ID_Tecnica_Interna": None,
            }))
        else:
            for tech, g in g_s.groupby("ID_Tecnica_Interna", dropna=True):
                rows.append(aggregate_group(g, {
                    "Escenario_Normalizado": scenario,
                    "ID_Tecnica_Interna": tech,
                }))

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def get_scenario_metric(agg_df: pd.DataFrame,
                        scenario: str,
                        metric: str,
                        tech: Optional[str] = None) -> Optional[float]:
    if agg_df.empty:
        return None
    g = agg_df[agg_df["Escenario_Normalizado"] == scenario]
    if tech is not None:
        g = g[g["ID_Tecnica_Interna"] == tech]
    if g.empty:
        return None
    return safe_float(g.iloc[0].get(f"{metric}_Media"))


def get_scenario_n(agg_df: pd.DataFrame, scenario: str, tech: Optional[str] = None) -> int:
    if agg_df.empty:
        return 0
    g = agg_df[agg_df["Escenario_Normalizado"] == scenario]
    if tech is not None:
        g = g[g["ID_Tecnica_Interna"] == tech]
    if g.empty:
        return 0
    value = safe_float(g.iloc[0].get("N_Repeticiones"))
    return int(value) if value is not None else 0


# ============================================================
# Clasificación heurística
# ============================================================

def classify_impact(cpu_impact_pp: object, ram_impact_mb: object) -> str:
    """
    Clasificación heurística solicitada.
    Usa impacto positivo. Si el delta bruto es negativo, se clasifica como 0.
    """
    cpu_raw = safe_float(cpu_impact_pp)
    ram_raw = safe_float(ram_impact_mb)

    if cpu_raw is None and ram_raw is None:
        return "No evaluado"

    cpu = max(0.0, cpu_raw or 0.0)
    ram = max(0.0, ram_raw or 0.0)

    if cpu > 15 or ram > 300:
        return "Alto"
    if cpu < 5 and ram < 100:
        return "Bajo"
    return "Medio"


# ============================================================
# Benchmark 4 escenarios
# ============================================================

def build_benchmark_4_scenarios(files_df: pd.DataFrame,
                                sample_interval_s: int,
                                tool_name: str) -> pd.DataFrame:
    agg = build_scenario_aggregates(files_df)

    techs = sorted(
        t for t in files_df["ID_Tecnica_Interna"].dropna().unique().tolist()
        if str(t).strip()
    )

    rows: List[Dict[str, object]] = []

    for tech in techs:
        # Valores medios por escenario.
        cpu_base_sin = get_scenario_metric(agg, "BASE_SIN_VR", "CPU_Total")
        cpu_base_con = get_scenario_metric(agg, "BASE_CON_VR", "CPU_Total")
        cpu_tec_sin = get_scenario_metric(agg, "TEC_SIN_VR", "CPU_Total", tech)
        cpu_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "CPU_Total", tech)

        ram_base_sin = get_scenario_metric(agg, "BASE_SIN_VR", "RAM_Available_MB")
        ram_base_con = get_scenario_metric(agg, "BASE_CON_VR", "RAM_Available_MB")
        ram_tec_sin = get_scenario_metric(agg, "TEC_SIN_VR", "RAM_Available_MB", tech)
        ram_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "RAM_Available_MB", tech)

        read_base_sin = get_scenario_metric(agg, "BASE_SIN_VR", "IO_Reads_sec")
        read_base_con = get_scenario_metric(agg, "BASE_CON_VR", "IO_Reads_sec")
        read_tec_sin = get_scenario_metric(agg, "TEC_SIN_VR", "IO_Reads_sec", tech)
        read_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "IO_Reads_sec", tech)

        write_base_sin = get_scenario_metric(agg, "BASE_SIN_VR", "IO_Writes_sec")
        write_base_con = get_scenario_metric(agg, "BASE_CON_VR", "IO_Writes_sec")
        write_tec_sin = get_scenario_metric(agg, "TEC_SIN_VR", "IO_Writes_sec", tech)
        write_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "IO_Writes_sec", tech)

        vr_cpu_base_con = get_scenario_metric(agg, "BASE_CON_VR", "VR_CPU")
        vr_cpu_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "VR_CPU", tech)
        vr_ram_base_con = get_scenario_metric(agg, "BASE_CON_VR", "VR_RAM_MB")
        vr_ram_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "VR_RAM_MB", tech)
        vr_read_base_con = get_scenario_metric(agg, "BASE_CON_VR", "VR_IO_Read_Bps")
        vr_read_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "VR_IO_Read_Bps", tech)
        vr_write_base_con = get_scenario_metric(agg, "BASE_CON_VR", "VR_IO_Write_Bps")
        vr_write_tec_con = get_scenario_metric(agg, "TEC_CON_VR", "VR_IO_Write_Bps", tech)

        vr_cpu_delta_tec_vs_base = delta(vr_cpu_tec_con, vr_cpu_base_con)
        vr_ram_delta_tec_vs_base = delta(vr_ram_tec_con, vr_ram_base_con)
        vr_read_delta_tec_vs_base = delta(vr_read_tec_con, vr_read_base_con)
        vr_write_delta_tec_vs_base = delta(vr_write_tec_con, vr_write_base_con)

        # Costes CPU en puntos porcentuales.
        cpu_coste_vr_reposo = delta(cpu_base_con, cpu_base_sin)
        cpu_coste_tec_sin = delta(cpu_tec_sin, cpu_base_sin)
        cpu_coste_tec_con = delta(cpu_tec_con, cpu_base_con)
        cpu_sobrecoste_vr = delta(cpu_coste_tec_con, cpu_coste_tec_sin)
        cpu_simple_con_vs_sin = delta(cpu_tec_con, cpu_tec_sin)

        # RAM impacto MB: base disponible - medida disponible.
        ram_coste_vr_reposo = ram_impact(ram_base_sin, ram_base_con)
        ram_coste_tec_sin = ram_impact(ram_base_sin, ram_tec_sin)
        ram_coste_tec_con = ram_impact(ram_base_con, ram_tec_con)
        ram_sobrecoste_vr = delta(ram_coste_tec_con, ram_coste_tec_sin)
        ram_simple_con_vs_sin = ram_impact(ram_tec_sin, ram_tec_con)

        # I/O deltas: medida - base.
        read_coste_vr_reposo = delta(read_base_con, read_base_sin)
        read_coste_tec_sin = delta(read_tec_sin, read_base_sin)
        read_coste_tec_con = delta(read_tec_con, read_base_con)
        read_sobrecoste_vr = delta(read_coste_tec_con, read_coste_tec_sin)
        read_simple_con_vs_sin = delta(read_tec_con, read_tec_sin)

        write_coste_vr_reposo = delta(write_base_con, write_base_sin)
        write_coste_tec_sin = delta(write_tec_sin, write_base_sin)
        write_coste_tec_con = delta(write_tec_con, write_base_con)
        write_sobrecoste_vr = delta(write_coste_tec_con, write_coste_tec_sin)
        write_simple_con_vs_sin = delta(write_tec_con, write_tec_sin)

        impacto_tec_con = classify_impact(cpu_coste_tec_con, ram_coste_tec_con)
        impacto_sobrecoste = classify_impact(cpu_sobrecoste_vr, ram_sobrecoste_vr)

        missing = []
        for s in ["BASE_SIN_VR", "BASE_CON_VR", "TEC_SIN_VR", "TEC_CON_VR"]:
            n = get_scenario_n(agg, s, tech if s.startswith("TEC") else None)
            if n == 0:
                missing.append(s)

        estado = "OK" if not missing else "Incompleto"
        obs = []
        if missing:
            obs.append("Faltan escenarios: " + ", ".join(missing))
        if cpu_sobrecoste_vr is not None and cpu_sobrecoste_vr < 0:
            obs.append("Sobrecoste CPU ajustado negativo; para clasificación se interpreta como 0.")
        if ram_sobrecoste_vr is not None and ram_sobrecoste_vr < 0:
            obs.append("Sobrecoste RAM ajustado negativo; para clasificación se interpreta como 0.")
        if not obs:
            obs.append("Comparativa ajustada calculada correctamente.")

        row = {
            "ID_Tecnica_Interna": tech,
            "Estado_Benchmark": estado,
            "Observaciones": " | ".join(obs),
            "Herramienta_Medicion": tool_name,
            "Intervalo_Muestreo_s": sample_interval_s,
            "N_BASE_SIN_VR": get_scenario_n(agg, "BASE_SIN_VR"),
            "N_BASE_CON_VR": get_scenario_n(agg, "BASE_CON_VR"),
            "N_TEC_SIN_VR": get_scenario_n(agg, "TEC_SIN_VR", tech),
            "N_TEC_CON_VR": get_scenario_n(agg, "TEC_CON_VR", tech),

            # Valores brutos por escenario.
            "CPU_BASE_SIN_VR_%": cpu_base_sin,
            "CPU_BASE_CON_VR_%": cpu_base_con,
            "CPU_TEC_SIN_VR_%": cpu_tec_sin,
            "CPU_TEC_CON_VR_%": cpu_tec_con,
            "RAM_Available_BASE_SIN_VR_MB": ram_base_sin,
            "RAM_Available_BASE_CON_VR_MB": ram_base_con,
            "RAM_Available_TEC_SIN_VR_MB": ram_tec_sin,
            "RAM_Available_TEC_CON_VR_MB": ram_tec_con,
            "IO_Read_BASE_SIN_VR": read_base_sin,
            "IO_Read_BASE_CON_VR": read_base_con,
            "IO_Read_TEC_SIN_VR": read_tec_sin,
            "IO_Read_TEC_CON_VR": read_tec_con,
            "IO_Write_BASE_SIN_VR": write_base_sin,
            "IO_Write_BASE_CON_VR": write_base_con,
            "IO_Write_TEC_SIN_VR": write_tec_sin,
            "IO_Write_TEC_CON_VR": write_tec_con,

            # Métricas proceso Velociraptor.
            "VR_CPU_BASE_CON_VR_%": vr_cpu_base_con,
            "VR_CPU_TEC_CON_VR_%": vr_cpu_tec_con,
            "VR_CPU_DELTA_TEC_vs_BASE_pp": vr_cpu_delta_tec_vs_base,
            "VR_RAM_BASE_CON_VR_MB": vr_ram_base_con,
            "VR_RAM_TEC_CON_VR_MB": vr_ram_tec_con,
            "VR_RAM_DELTA_TEC_vs_BASE_MB": vr_ram_delta_tec_vs_base,
            "VR_IO_Read_BASE_CON_VR_Bps": vr_read_base_con,
            "VR_IO_Read_TEC_CON_VR_Bps": vr_read_tec_con,
            "VR_IO_Read_DELTA_TEC_vs_BASE_Bps": vr_read_delta_tec_vs_base,
            "VR_IO_Write_BASE_CON_VR_Bps": vr_write_base_con,
            "VR_IO_Write_TEC_CON_VR_Bps": vr_write_tec_con,
            "VR_IO_Write_DELTA_TEC_vs_BASE_Bps": vr_write_delta_tec_vs_base,

            # Costes principales.
            "Coste_VR_Reposo_CPU_pp": cpu_coste_vr_reposo,
            "Coste_Tecnica_SIN_VR_CPU_pp": cpu_coste_tec_sin,
            "Coste_Tecnica_CON_VR_CPU_pp": cpu_coste_tec_con,
            "Sobrecoste_VR_Durante_Tecnica_CPU_pp": cpu_sobrecoste_vr,
            "Comparacion_Simple_TEC_CON_vs_SIN_CPU_pp": cpu_simple_con_vs_sin,

            "Coste_VR_Reposo_RAM_MB": ram_coste_vr_reposo,
            "Coste_Tecnica_SIN_VR_RAM_MB": ram_coste_tec_sin,
            "Coste_Tecnica_CON_VR_RAM_MB": ram_coste_tec_con,
            "Sobrecoste_VR_Durante_Tecnica_RAM_MB": ram_sobrecoste_vr,
            "Comparacion_Simple_TEC_CON_vs_SIN_RAM_MB": ram_simple_con_vs_sin,

            "Coste_VR_Reposo_IO_Read": read_coste_vr_reposo,
            "Coste_Tecnica_SIN_VR_IO_Read": read_coste_tec_sin,
            "Coste_Tecnica_CON_VR_IO_Read": read_coste_tec_con,
            "Sobrecoste_VR_Durante_Tecnica_IO_Read": read_sobrecoste_vr,
            "Comparacion_Simple_TEC_CON_vs_SIN_IO_Read": read_simple_con_vs_sin,

            "Coste_VR_Reposo_IO_Write": write_coste_vr_reposo,
            "Coste_Tecnica_SIN_VR_IO_Write": write_coste_tec_sin,
            "Coste_Tecnica_CON_VR_IO_Write": write_coste_tec_con,
            "Sobrecoste_VR_Durante_Tecnica_IO_Write": write_sobrecoste_vr,
            "Comparacion_Simple_TEC_CON_vs_SIN_IO_Write": write_simple_con_vs_sin,

            "Impacto_Tecnica_CON_VR": impacto_tec_con,
            "Impacto_Sobrecoste_VR": impacto_sobrecoste,
        }
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows).sort_values(by=["ID_Tecnica_Interna"]).reset_index(drop=True)

    # Redondeo técnico para legibilidad.
    for col in df.columns:
        if df[col].dtype.kind in "fc":
            df[col] = df[col].map(lambda x: round_or_none(x, 3))
    return df


def build_resumen_tecnicas(benchmark_df: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla final para la memoria.
    Incluye:
    - impacto de la técnica sin VR;
    - impacto de la técnica con VR;
    - sobrecoste ajustado atribuible a Velociraptor;
    - comparación simple CON_VR vs SIN_VR;
    - métricas directas del proceso Velociraptor.
    """
    cols = [
        "ID_Tecnica_Interna",
        "CPU_Coste_VR_Reposo_pp",
        "CPU_Tecnica_SIN_VR_pp",
        "CPU_Tecnica_CON_VR_pp",
        "CPU_Sobrecoste_VR_pp",
        "CPU_Diff_Simple_CON_vs_SIN_pp",
        "RAM_Coste_VR_Reposo_MB",
        "RAM_Tecnica_SIN_VR_MB",
        "RAM_Tecnica_CON_VR_MB",
        "RAM_Sobrecoste_VR_MB",
        "RAM_Diff_Simple_CON_vs_SIN_MB",
        "IO_Read_Tecnica_SIN_VR",
        "IO_Read_Tecnica_CON_VR",
        "IO_Read_Sobrecoste_VR",
        "IO_Read_Diff_Simple_CON_vs_SIN",
        "IO_Write_Tecnica_SIN_VR",
        "IO_Write_Tecnica_CON_VR",
        "IO_Write_Sobrecoste_VR",
        "IO_Write_Diff_Simple_CON_vs_SIN",
        "VR_CPU_BASE_CON_VR_%",
        "VR_CPU_TEC_CON_VR_%",
        "VR_CPU_DELTA_TEC_vs_BASE_pp",
        "VR_RAM_BASE_CON_VR_MB",
        "VR_RAM_TEC_CON_VR_MB",
        "VR_RAM_DELTA_TEC_vs_BASE_MB",
        "VR_IO_Read_BASE_CON_VR_Bps",
        "VR_IO_Read_TEC_CON_VR_Bps",
        "VR_IO_Read_DELTA_TEC_vs_BASE_Bps",
        "VR_IO_Write_BASE_CON_VR_Bps",
        "VR_IO_Write_TEC_CON_VR_Bps",
        "VR_IO_Write_DELTA_TEC_vs_BASE_Bps",
        "Impacto_Tecnica_CON_VR",
        "Impacto_Rendimiento",
        "Observaciones",
    ]
    if benchmark_df.empty:
        return pd.DataFrame(columns=cols)

    out = pd.DataFrame({
        "ID_Tecnica_Interna": benchmark_df["ID_Tecnica_Interna"],
        "CPU_Coste_VR_Reposo_pp": benchmark_df["Coste_VR_Reposo_CPU_pp"],
        "CPU_Tecnica_SIN_VR_pp": benchmark_df["Coste_Tecnica_SIN_VR_CPU_pp"],
        "CPU_Tecnica_CON_VR_pp": benchmark_df["Coste_Tecnica_CON_VR_CPU_pp"],
        "CPU_Sobrecoste_VR_pp": benchmark_df["Sobrecoste_VR_Durante_Tecnica_CPU_pp"],
        "CPU_Diff_Simple_CON_vs_SIN_pp": benchmark_df["Comparacion_Simple_TEC_CON_vs_SIN_CPU_pp"],
        "RAM_Coste_VR_Reposo_MB": benchmark_df["Coste_VR_Reposo_RAM_MB"],
        "RAM_Tecnica_SIN_VR_MB": benchmark_df["Coste_Tecnica_SIN_VR_RAM_MB"],
        "RAM_Tecnica_CON_VR_MB": benchmark_df["Coste_Tecnica_CON_VR_RAM_MB"],
        "RAM_Sobrecoste_VR_MB": benchmark_df["Sobrecoste_VR_Durante_Tecnica_RAM_MB"],
        "RAM_Diff_Simple_CON_vs_SIN_MB": benchmark_df["Comparacion_Simple_TEC_CON_vs_SIN_RAM_MB"],
        "IO_Read_Tecnica_SIN_VR": benchmark_df["Coste_Tecnica_SIN_VR_IO_Read"],
        "IO_Read_Tecnica_CON_VR": benchmark_df["Coste_Tecnica_CON_VR_IO_Read"],
        "IO_Read_Sobrecoste_VR": benchmark_df["Sobrecoste_VR_Durante_Tecnica_IO_Read"],
        "IO_Read_Diff_Simple_CON_vs_SIN": benchmark_df["Comparacion_Simple_TEC_CON_vs_SIN_IO_Read"],
        "IO_Write_Tecnica_SIN_VR": benchmark_df["Coste_Tecnica_SIN_VR_IO_Write"],
        "IO_Write_Tecnica_CON_VR": benchmark_df["Coste_Tecnica_CON_VR_IO_Write"],
        "IO_Write_Sobrecoste_VR": benchmark_df["Sobrecoste_VR_Durante_Tecnica_IO_Write"],
        "IO_Write_Diff_Simple_CON_vs_SIN": benchmark_df["Comparacion_Simple_TEC_CON_vs_SIN_IO_Write"],
        "VR_CPU_BASE_CON_VR_%": benchmark_df.get("VR_CPU_BASE_CON_VR_%"),
        "VR_CPU_TEC_CON_VR_%": benchmark_df.get("VR_CPU_TEC_CON_VR_%"),
        "VR_CPU_DELTA_TEC_vs_BASE_pp": benchmark_df.get("VR_CPU_DELTA_TEC_vs_BASE_pp"),
        "VR_RAM_BASE_CON_VR_MB": benchmark_df.get("VR_RAM_BASE_CON_VR_MB"),
        "VR_RAM_TEC_CON_VR_MB": benchmark_df.get("VR_RAM_TEC_CON_VR_MB"),
        "VR_RAM_DELTA_TEC_vs_BASE_MB": benchmark_df.get("VR_RAM_DELTA_TEC_vs_BASE_MB"),
        "VR_IO_Read_BASE_CON_VR_Bps": benchmark_df.get("VR_IO_Read_BASE_CON_VR_Bps"),
        "VR_IO_Read_TEC_CON_VR_Bps": benchmark_df.get("VR_IO_Read_TEC_CON_VR_Bps"),
        "VR_IO_Read_DELTA_TEC_vs_BASE_Bps": benchmark_df.get("VR_IO_Read_DELTA_TEC_vs_BASE_Bps"),
        "VR_IO_Write_BASE_CON_VR_Bps": benchmark_df.get("VR_IO_Write_BASE_CON_VR_Bps"),
        "VR_IO_Write_TEC_CON_VR_Bps": benchmark_df.get("VR_IO_Write_TEC_CON_VR_Bps"),
        "VR_IO_Write_DELTA_TEC_vs_BASE_Bps": benchmark_df.get("VR_IO_Write_DELTA_TEC_vs_BASE_Bps"),
        "Impacto_Tecnica_CON_VR": benchmark_df["Impacto_Tecnica_CON_VR"],
        "Impacto_Rendimiento": benchmark_df["Impacto_Sobrecoste_VR"],
        "Observaciones": benchmark_df["Observaciones"],
    })
    return out[cols].reset_index(drop=True)


def build_sobrecoste_global_vr(benchmark_df: pd.DataFrame) -> pd.DataFrame:
    """
    Estimación agregada para discutir en conclusiones.
    No sustituye el análisis por técnica: resume las métricas principales.
    """
    cols = [
        "Grupo", "Metrica", "Unidad", "Media", "Mediana", "STD", "Min", "Max", "N", "Interpretacion"
    ]
    if benchmark_df.empty:
        return pd.DataFrame(columns=cols)

    metrics = [
        ("Coste VR en reposo", "Coste_VR_Reposo_CPU_pp", "pp", "CPU adicional del sistema al activar Velociraptor en reposo."),
        ("Coste VR en reposo", "Coste_VR_Reposo_RAM_MB", "MB", "RAM adicional interpretada desde menor memoria disponible al activar Velociraptor."),
        ("Coste VR en reposo", "Coste_VR_Reposo_IO_Read", "lecturas/s", "Diferencia de lecturas de disco en reposo con VR frente a sin VR."),
        ("Coste VR en reposo", "Coste_VR_Reposo_IO_Write", "escrituras/s", "Diferencia de escrituras de disco en reposo con VR frente a sin VR."),
        ("Sobrecoste ajustado VR", "Sobrecoste_VR_Durante_Tecnica_CPU_pp", "pp", "Sobrecoste CPU atribuible a Velociraptor durante la técnica."),
        ("Sobrecoste ajustado VR", "Sobrecoste_VR_Durante_Tecnica_RAM_MB", "MB", "Sobrecoste RAM atribuible a Velociraptor durante la técnica."),
        ("Sobrecoste ajustado VR", "Sobrecoste_VR_Durante_Tecnica_IO_Read", "lecturas/s", "Sobrecoste ajustado de lecturas de disco durante la técnica."),
        ("Sobrecoste ajustado VR", "Sobrecoste_VR_Durante_Tecnica_IO_Write", "escrituras/s", "Sobrecoste ajustado de escrituras de disco durante la técnica."),
        ("Comparación simple", "Comparacion_Simple_TEC_CON_vs_SIN_CPU_pp", "pp", "Diferencia directa TEC_CON_VR - TEC_SIN_VR; métrica secundaria."),
        ("Comparación simple", "Comparacion_Simple_TEC_CON_vs_SIN_RAM_MB", "MB", "Diferencia directa en RAM interpretada; métrica secundaria."),
        ("Proceso Velociraptor", "VR_CPU_DELTA_TEC_vs_BASE_pp", "pp", "Cambio de CPU del proceso Velociraptor durante técnica frente a reposo con VR."),
        ("Proceso Velociraptor", "VR_RAM_DELTA_TEC_vs_BASE_MB", "MB", "Cambio de RAM del proceso Velociraptor durante técnica frente a reposo con VR."),
        ("Proceso Velociraptor", "VR_IO_Read_DELTA_TEC_vs_BASE_Bps", "B/s", "Cambio de lecturas del proceso Velociraptor durante técnica frente a reposo con VR."),
        ("Proceso Velociraptor", "VR_IO_Write_DELTA_TEC_vs_BASE_Bps", "B/s", "Cambio de escrituras del proceso Velociraptor durante técnica frente a reposo con VR."),
    ]

    rows: List[Dict[str, object]] = []
    for group, metric, unit, interpretation in metrics:
        if metric not in benchmark_df.columns:
            rows.append({
                "Grupo": group, "Metrica": metric, "Unidad": unit,
                "Media": None, "Mediana": None, "STD": None, "Min": None, "Max": None, "N": 0,
                "Interpretacion": interpretation,
            })
            continue
        s = pd.to_numeric(benchmark_df[metric], errors="coerce").dropna()
        if s.empty:
            rows.append({
                "Grupo": group, "Metrica": metric, "Unidad": unit,
                "Media": None, "Mediana": None, "STD": None, "Min": None, "Max": None, "N": 0,
                "Interpretacion": interpretation,
            })
        else:
            rows.append({
                "Grupo": group,
                "Metrica": metric,
                "Unidad": unit,
                "Media": round_or_none(s.mean()),
                "Mediana": round_or_none(s.median()),
                "STD": round_or_none(s.std(ddof=1) if len(s) > 1 else 0.0),
                "Min": round_or_none(s.min()),
                "Max": round_or_none(s.max()),
                "N": int(s.count()),
                "Interpretacion": interpretation,
            })

    return pd.DataFrame(rows, columns=cols)


# ============================================================
# Excel
# ============================================================

def autofit_columns(worksheet, dataframe: pd.DataFrame, max_width: int = 48) -> None:
    if dataframe is None or dataframe.empty:
        return
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].head(300).fillna("").astype(str).tolist()
        max_len = max([len(str(col))] + [len(v) for v in values])
        width = min(max(max_len + 2, 10), max_width)
        worksheet.set_column(idx, idx, width)


def apply_basic_sheet_format(writer, sheet_name: str, df: pd.DataFrame, workbook) -> None:
    ws = writer.sheets[sheet_name]
    fmt_header = workbook.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#1F4E78",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
    })
    fmt_num = workbook.add_format({"num_format": "0.000"})
    fmt_int = workbook.add_format({"num_format": "0"})
    fmt_text_wrap = workbook.add_format({"text_wrap": True, "valign": "top"})

    if df is None:
        return

    if len(df.columns) > 0:
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)

    for c, col in enumerate(df.columns):
        ws.write(0, c, col, fmt_header)

    autofit_columns(ws, df)

    for c, col in enumerate(df.columns):
        lower = str(col).lower()
        if any(k in lower for k in [
            "cpu", "ram", "io_", "duracion", "media", "std", "delta", "coste", "sobrecoste", "%", "mb", "bps"
        ]):
            ws.set_column(c, c, min(max(len(str(col)) + 2, 12), 22), fmt_num)
        elif any(k in lower for k in ["n_", "repeticion", "filas", "columnas", "samples", "muestras", "intervalo"]):
            ws.set_column(c, c, min(max(len(str(col)) + 2, 10), 18), fmt_int)
        elif any(k in lower for k in ["observaciones", "advertencias", "archivos", "ruta", "error"]):
            ws.set_column(c, c, 55, fmt_text_wrap)

    # Ocultar columnas de trazabilidad de cabeceras para no ensuciar, pero mantenerlas disponibles.
    for c, col in enumerate(df.columns):
        if str(col).startswith("Cols_"):
            ws.set_column(c, c, None, None, {"hidden": True})


def add_chart(workbook,
              worksheet,
              sheet_name: str,
              df: pd.DataFrame,
              title: str,
              category_col: str,
              value_cols: Sequence[str],
              cell: str,
              y_axis: str) -> None:
    if df is None or df.empty:
        return
    col_index = {col: idx for idx, col in enumerate(df.columns)}
    if category_col not in col_index:
        return
    if any(col not in col_index for col in value_cols):
        return

    chart = workbook.add_chart({"type": "column"})
    for value_col in value_cols:
        chart.add_series({
            "name": [sheet_name, 0, col_index[value_col]],
            "categories": [sheet_name, 1, col_index[category_col], len(df), col_index[category_col]],
            "values": [sheet_name, 1, col_index[value_col], len(df), col_index[value_col]],
        })
    chart.set_title({"name": title})
    chart.set_x_axis({"name": "Técnica"})
    chart.set_y_axis({"name": y_axis})
    chart.set_legend({"position": "bottom"})
    chart.set_style(10)
    worksheet.insert_chart(cell, chart, {"x_scale": 1.35, "y_scale": 1.18})


def write_excel(output_xlsx: Path,
                files_df: pd.DataFrame,
                qc_df: pd.DataFrame,
                baselines_df: pd.DataFrame,
                benchmark_df: pd.DataFrame,
                resumen_df: pd.DataFrame,
                global_df: pd.DataFrame,
                errors_df: pd.DataFrame) -> None:
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_xlsx, engine="xlsxwriter") as writer:
        sheets = {
            "00_Metricas_Por_Archivo": files_df,
            "01_Control_Calidad": qc_df,
            "02_Baselines": baselines_df,
            "03_Benchmark_4_Escenarios": benchmark_df,
            "04_Resumen_Tecnicas": resumen_df,
            "06_Sobrecoste_Global_VR": global_df,
        }

        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        if errors_df is not None and not errors_df.empty:
            errors_df.to_excel(writer, sheet_name="99_Errores", index=False)

        workbook = writer.book

        for sheet_name, df in sheets.items():
            apply_basic_sheet_format(writer, sheet_name, df, workbook)

        if errors_df is not None and not errors_df.empty:
            apply_basic_sheet_format(writer, "99_Errores", errors_df, workbook)

        # Hoja de gráficos.
        ws_graph = workbook.add_worksheet("05_Graficos")
        writer.sheets["05_Graficos"] = ws_graph

        fmt_title = workbook.add_format({"bold": True, "font_size": 15})
        fmt_section = workbook.add_format({"bold": True, "font_size": 12})
        fmt_note = workbook.add_format({"text_wrap": True, "valign": "top"})

        ws_graph.write("A1", "Gráficos automáticos - Benchmark Velociraptor 4 escenarios", fmt_title)
        ws_graph.write(
            "A3",
            "Comparación principal: sobrecoste ajustado = "
            "(TEC_CON_VR - BASE_CON_VR) - (TEC_SIN_VR - BASE_SIN_VR). "
            "La comparación simple TEC_CON_VR - TEC_SIN_VR se mantiene como métrica secundaria.",
            fmt_note,
        )
        ws_graph.set_column("A:A", 95)

        ws_graph.write("A5", "1. Valores brutos por escenario", fmt_section)
        add_chart(
            workbook, ws_graph, "03_Benchmark_4_Escenarios", benchmark_df,
            "CPU bruta: 4 escenarios",
            "ID_Tecnica_Interna",
            ["CPU_BASE_SIN_VR_%", "CPU_BASE_CON_VR_%", "CPU_TEC_SIN_VR_%", "CPU_TEC_CON_VR_%"],
            "A6", "CPU media (%)",
        )
        add_chart(
            workbook, ws_graph, "03_Benchmark_4_Escenarios", benchmark_df,
            "RAM disponible: 4 escenarios",
            "ID_Tecnica_Interna",
            ["RAM_Available_BASE_SIN_VR_MB", "RAM_Available_BASE_CON_VR_MB", "RAM_Available_TEC_SIN_VR_MB", "RAM_Available_TEC_CON_VR_MB"],
            "J6", "RAM disponible (MB)",
        )
        add_chart(
            workbook, ws_graph, "03_Benchmark_4_Escenarios", benchmark_df,
            "I/O reads bruto: 4 escenarios",
            "ID_Tecnica_Interna",
            ["IO_Read_BASE_SIN_VR", "IO_Read_BASE_CON_VR", "IO_Read_TEC_SIN_VR", "IO_Read_TEC_CON_VR"],
            "A24", "Disk Reads/sec",
        )
        add_chart(
            workbook, ws_graph, "03_Benchmark_4_Escenarios", benchmark_df,
            "I/O writes bruto: 4 escenarios",
            "ID_Tecnica_Interna",
            ["IO_Write_BASE_SIN_VR", "IO_Write_BASE_CON_VR", "IO_Write_TEC_SIN_VR", "IO_Write_TEC_CON_VR"],
            "J24", "Disk Writes/sec",
        )

        ws_graph.write("A42", "2. Coste de técnica y sobrecoste ajustado", fmt_section)
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "CPU: técnica SIN_VR vs CON_VR vs sobrecoste ajustado",
            "ID_Tecnica_Interna",
            ["CPU_Tecnica_SIN_VR_pp", "CPU_Tecnica_CON_VR_pp", "CPU_Sobrecoste_VR_pp"],
            "A43", "CPU delta (puntos porcentuales)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "RAM: técnica SIN_VR vs CON_VR vs sobrecoste ajustado",
            "ID_Tecnica_Interna",
            ["RAM_Tecnica_SIN_VR_MB", "RAM_Tecnica_CON_VR_MB", "RAM_Sobrecoste_VR_MB"],
            "J43", "RAM impacto (MB)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "I/O reads: técnica SIN_VR vs CON_VR vs sobrecoste ajustado",
            "ID_Tecnica_Interna",
            ["IO_Read_Tecnica_SIN_VR", "IO_Read_Tecnica_CON_VR", "IO_Read_Sobrecoste_VR"],
            "A61", "Disk Reads/sec delta",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "I/O writes: técnica SIN_VR vs CON_VR vs sobrecoste ajustado",
            "ID_Tecnica_Interna",
            ["IO_Write_Tecnica_SIN_VR", "IO_Write_Tecnica_CON_VR", "IO_Write_Sobrecoste_VR"],
            "J61", "Disk Writes/sec delta",
        )

        ws_graph.write("A79", "3. Sobrecoste ajustado frente a comparación simple", fmt_section)
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "CPU: sobrecoste ajustado vs comparación simple",
            "ID_Tecnica_Interna",
            ["CPU_Sobrecoste_VR_pp", "CPU_Diff_Simple_CON_vs_SIN_pp"],
            "A80", "CPU delta (puntos porcentuales)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "RAM: sobrecoste ajustado vs comparación simple",
            "ID_Tecnica_Interna",
            ["RAM_Sobrecoste_VR_MB", "RAM_Diff_Simple_CON_vs_SIN_MB"],
            "J80", "RAM impacto (MB)",
        )

        ws_graph.write("A98", "4. Coste estimado de Velociraptor en reposo", fmt_section)
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Coste CPU de Velociraptor en reposo",
            "ID_Tecnica_Interna",
            ["CPU_Coste_VR_Reposo_pp"],
            "A99", "CPU delta (pp)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Coste RAM de Velociraptor en reposo",
            "ID_Tecnica_Interna",
            ["RAM_Coste_VR_Reposo_MB"],
            "J99", "RAM impacto (MB)",
        )

        ws_graph.write("A117", "5. Métricas directas del proceso Velociraptor", fmt_section)
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Proceso Velociraptor: CPU reposo vs técnica",
            "ID_Tecnica_Interna",
            ["VR_CPU_BASE_CON_VR_%", "VR_CPU_TEC_CON_VR_%", "VR_CPU_DELTA_TEC_vs_BASE_pp"],
            "A118", "CPU proceso VR (%)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Proceso Velociraptor: RAM reposo vs técnica",
            "ID_Tecnica_Interna",
            ["VR_RAM_BASE_CON_VR_MB", "VR_RAM_TEC_CON_VR_MB", "VR_RAM_DELTA_TEC_vs_BASE_MB"],
            "J118", "RAM proceso VR (MB)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Proceso Velociraptor: I/O reads reposo vs técnica",
            "ID_Tecnica_Interna",
            ["VR_IO_Read_BASE_CON_VR_Bps", "VR_IO_Read_TEC_CON_VR_Bps", "VR_IO_Read_DELTA_TEC_vs_BASE_Bps"],
            "A136", "VR IO Read (B/s)",
        )
        add_chart(
            workbook, ws_graph, "04_Resumen_Tecnicas", resumen_df,
            "Proceso Velociraptor: I/O writes reposo vs técnica",
            "ID_Tecnica_Interna",
            ["VR_IO_Write_BASE_CON_VR_Bps", "VR_IO_Write_TEC_CON_VR_Bps", "VR_IO_Write_DELTA_TEC_vs_BASE_Bps"],
            "J136", "VR IO Write (B/s)",
        )

        # Hoja 06 también queda lista para usar como tabla agregada de conclusiones.


# ============================================================
# Main
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Procesa CSV de logman y genera Excel para benchmark Velociraptor 4 escenarios."
    )
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR, help="Carpeta con CSV de logman")
    parser.add_argument("--output-xlsx", type=Path, default=DEFAULT_OUTPUT_XLSX, help="Ruta del Excel generado")
    parser.add_argument("--sample-interval", type=int, default=DEFAULT_SAMPLE_INTERVAL_S, help="Intervalo de muestreo en segundos")
    parser.add_argument("--tool-name", type=str, default=DEFAULT_TOOL_NAME, help="Nombre de la herramienta de medición")
    parser.add_argument("--process-name", type=str, default=DEFAULT_PROCESS_NAME, help="Nombre del proceso Velociraptor en PerfMon")
    parser.add_argument("--min-duration", type=int, default=DEFAULT_MIN_DURATION_S, help="Duración mínima esperada por CSV en segundos")
    args = parser.parse_args()

    input_dir: Path = args.input_dir
    output_xlsx: Path = args.output_xlsx

    if not input_dir.exists():
        raise FileNotFoundError(f"No existe la carpeta de entrada: {input_dir}")

    csv_files = sorted(input_dir.glob("*.csv"))
    if not csv_files:
        raise FileNotFoundError(f"No se encontraron CSV en: {input_dir}")

    rows: List[Dict[str, object]] = []
    errors: List[Dict[str, object]] = []

    for csv_path in csv_files:
        try:
            rows.append(extract_metrics_from_file(csv_path, process_name=args.process_name))
        except Exception as exc:
            errors.append({
                "Archivo": csv_path.name,
                "Error": str(exc),
                "Causa_Probable": "CSV corrupto, codificación no soportada, separador inesperado o cabeceras no compatibles.",
            })

    if not rows:
        raise RuntimeError("No se pudo procesar ningún CSV correctamente.")

    files_df = pd.DataFrame(rows).sort_values(by=["Archivo"]).reset_index(drop=True)
    errors_df = pd.DataFrame(errors, columns=["Archivo", "Error", "Causa_Probable"])

    qc_df = build_quality_control(files_df, errors_df, min_duration_s=args.min_duration)
    baselines_df = build_baselines(files_df)
    benchmark_df = build_benchmark_4_scenarios(files_df, sample_interval_s=args.sample_interval, tool_name=args.tool_name)
    resumen_df = build_resumen_tecnicas(benchmark_df)
    global_df = build_sobrecoste_global_vr(benchmark_df)

    write_excel(
        output_xlsx=output_xlsx,
        files_df=files_df,
        qc_df=qc_df,
        baselines_df=baselines_df,
        benchmark_df=benchmark_df,
        resumen_df=resumen_df,
        global_df=global_df,
        errors_df=errors_df,
    )

    print(f"[OK] Excel generado: {output_xlsx}")
    print(f"[OK] CSV procesados correctamente: {len(files_df)}")
    print(f"[OK] Técnicas en 03_Benchmark_4_Escenarios: {len(benchmark_df)}")
    print(f"[OK] Filas en 04_Resumen_Tecnicas: {len(resumen_df)}")
    print(f"[OK] Filas en 06_Sobrecoste_Global_VR: {len(global_df)}")
    if not errors_df.empty:
        print(f"[WARN] CSV con error: {len(errors_df)}. Revisar hoja 99_Errores.")


if __name__ == "__main__":
    main()
