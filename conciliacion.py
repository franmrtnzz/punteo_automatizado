#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║     SISTEMA DE CONCILIACIÓN BANCARIA AUTOMATIZADA                          ║
║     Punteo automático: Extracto Bancario ↔ Libro Mayor (572.1)            ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  POLÍTICA: MÁXIMA PRECISIÓN, CERO FALSOS POSITIVOS                        ║
║  → Ante cualquier duda → PENDIENTE DE REVISIÓN                            ║
║  → NUNCA marcar como conciliado si hay incertidumbre                       ║
╚══════════════════════════════════════════════════════════════════════════════╝

Uso:
    python3 conciliacion.py

    Coloca los archivos Excel del banco y del mayor en la misma carpeta
    que este script. El sistema los detectará automáticamente.

    Archivos esperados:
    - Un archivo con "BANCO" en el nombre (.xls o .xlsx)
    - Un archivo con "MAYOR" en el nombre (.xls o .xlsx)
"""

import pandas as pd
import numpy as np
from pathlib import Path
from itertools import combinations
from datetime import datetime, timedelta
import warnings
import sys
import os
import unicodedata

warnings.filterwarnings('ignore')

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN (ajustar según empresa)
# ═══════════════════════════════════════════════════════════════════════════════
CONFIG = {
    # --- Formato de fechas (sugerencia para auto-detección) ---
    'mayor_date_format': '%d/%m/%Y',

    # --- Parámetros de matching (conservadores) ---
    'ventana_dias_pass2': 3,    # ±3 días para Pass 2
    'ventana_dias_pass3': 15,   # ±15 días para Pass 3
    'ventana_dias_pass4': 60,   # ±60 días para Pass 4 (importes únicos)
    'ventana_dias_duplicados': 10,  # ±10 días para duplicados (Pass 5)
    'ventana_dias_split': 5,    # ±5 días para detección de splits (conservador)
    'max_split_parts': 4,       # Máximo de partes en un split (1:N)

    # --- Período de conciliación ---
    'fecha_inicio': '2025-01-01',
    'fecha_fin': '2025-12-31',
}


# ═══════════════════════════════════════════════════════════════════════════════
#  DETECCIÓN AUTOMÁTICA DE LAYOUT  (v2 — robusta)
# ═══════════════════════════════════════════════════════════════════════════════

def _normalize_text(text):
    """Normaliza texto: minúsculas, sin acentos, espacios simples."""
    if pd.isna(text):
        return ''
    text = str(text).strip()
    text = unicodedata.normalize('NFD', text)
    text = ''.join(c for c in text if unicodedata.category(c) != 'Mn')
    text = text.lower().replace('_', ' ').replace('-', ' ').replace('.', ' ')
    return ' '.join(text.split())


# ── Diccionarios de keywords (ampliados) ──────────────────────────────────────

_KW_FECHA = [
    'fecha', 'date', 'fec', 'fecha valor', 'fecha operacion',
    'fecha contable', 'f valor', 'f operacion', 'value date',
]
_KW_IMPORTE = [
    'importe', 'cantidad', 'amount', 'euros', 'monto', 'valor',
    'total', 'sum', 'cuantia',
]
_KW_IMPORTE_EXCL = [
    'debe', 'haber', 'saldo', 'balance', 'acumulado', 'red', 'reducido',
    'anterior', 'neto',
]
_KW_DEBE = ['debe', 'cargo', 'cargos', 'debit', 'debito']
_KW_HABER = ['haber', 'abono', 'abonos', 'credit', 'credito']
_KW_SALDO = ['saldo', 'balance', 'acumulado', 'saldo contable']
_KW_MOVIMIENTO = [
    'movimiento', 'concepto', 'descripcion', 'detalle', 'beneficiario',
    'titular', 'nombre', 'ordenante', 'pagador', 'remitente',
    'destinatario', 'librador', 'receptor', 'proveedor',
    'description', 'narrative', 'payee', 'transaction',
]
_KW_MAS_DATOS = [
    'mas datos', 'datos adicional', 'observacion', 'referencia',
    'informacion', 'adicional', 'complemento', 'nota', 'detalle adicional',
    'ref', 'additional', 'remarks', 'memo',
]
_KW_DOCUMENTO = [
    'documento', 'doc', 'asiento', 'num doc', 'numero', 'num asiento',
    'n asiento', 'no doc', 'voucher', 'entry', 'nº', 'num',
]
_KW_CONCEPTO = [
    'concepto', 'descripcion', 'detalle', 'texto', 'glosa', 'narrative',
]
_KW_CONTRAPARTIDA = [
    'contrapartida', 'cuenta contra', 'contra', 'cuenta', 'cta',
    'subcuenta', 'auxiliar',
]
_KW_NETO = ['importe neto', 'neto', 'net', 'importe net']
_KW_RED  = ['importe red', 'reducido', 'importe reducido']
_KW_MARCA = ['marca', 'check', 'punteo', 'conciliado', 'ok']

# Uniones para detección de cabeceras
_BANCO_HEADER_KW = list(set(
    _KW_FECHA + _KW_IMPORTE + _KW_DEBE + _KW_HABER +
    _KW_MOVIMIENTO + _KW_MAS_DATOS
))
_MAYOR_HEADER_KW = list(set(
    _KW_FECHA + _KW_IMPORTE + _KW_DEBE + _KW_HABER + _KW_SALDO +
    _KW_DOCUMENTO + _KW_CONCEPTO + _KW_CONTRAPARTIDA +
    _KW_NETO + _KW_RED + _KW_MARCA
))


# ── Helpers internos ──────────────────────────────────────────────────────────

def _to_numeric_robust(series):
    """
    Convierte a numérico detectando automáticamente separador decimal.
    Soporta '1.234,56' (coma decimal) y '1,234.56' (punto decimal).
    """
    # Intento directo
    result = pd.to_numeric(series, errors='coerce')
    pct_ok = result.notna().sum() / max(len(series), 1)
    if pct_ok > 0.3:
        return result

    # Probablemente coma como decimal: "1.234,56" → "1234.56"
    try:
        cleaned = (
            series.astype(str)
            .str.strip()
            .str.replace(r'[^\d,.\-+]', '', regex=True)   # quitar €, espacios…
            .str.replace('.', '', regex=False)              # quitar separador miles
            .str.replace(',', '.', regex=False)             # coma → punto decimal
        )
        alt = pd.to_numeric(cleaned, errors='coerce')
        if alt.notna().sum() > result.notna().sum():
            return alt
    except Exception:
        pass

    return result


def _parse_dates(series):
    """Intenta parsear fechas con múltiples estrategias de fallback."""
    # 1. Conversión directa (dayfirst para formato europeo)
    result = pd.to_datetime(series, dayfirst=True, errors='coerce')
    if result.notna().sum() > len(series) * 0.3:
        return result

    # 2. Formato configurado
    fmt = CONFIG.get('mayor_date_format')
    if fmt:
        alt = pd.to_datetime(series, format=fmt, errors='coerce')
        if alt.notna().sum() > result.notna().sum():
            result = alt

    # 3. Formatos comunes
    for f in ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d.%m.%Y',
              '%d/%m/%y', '%d-%m-%y', '%Y%m%d', '%m/%d/%Y']:
        alt = pd.to_datetime(series, format=f, errors='coerce')
        if alt.notna().sum() > result.notna().sum():
            result = alt

    return result


def _is_numeric_col(series, threshold=0.3):
    """Devuelve True si al menos *threshold* de la serie es numérica."""
    num = pd.to_numeric(series, errors='coerce')
    return num.notna().sum() / max(len(series), 1) > threshold


def _is_date_col(series, threshold=0.3):
    """Devuelve True si al menos *threshold* de la serie parsea como fecha."""
    dates = pd.to_datetime(series, dayfirst=True, errors='coerce')
    return dates.notna().sum() / max(len(series), 1) > threshold


def _get_best_sheet(filepath, keywords, max_scan=30):
    """
    Si el Excel tiene varias hojas, devuelve el nombre de la hoja
    con mayor puntuación de keywords en sus primeras filas.
    Si solo tiene una hoja, la devuelve directamente.
    """
    try:
        xl = pd.ExcelFile(filepath)
    except Exception:
        return 0  # fallback

    sheet_names = xl.sheet_names
    if len(sheet_names) == 1:
        return sheet_names[0]

    best_sheet = sheet_names[0]
    best_score = -1

    for sn in sheet_names:
        try:
            df_raw = pd.read_excel(filepath, sheet_name=sn,
                                   header=None, nrows=max_scan)
        except Exception:
            continue

        score = 0
        for idx in range(len(df_raw)):
            row_vals = df_raw.iloc[idx].tolist()
            score += sum(
                1 for v in row_vals
                if _normalize_text(v)
                and any(kw in _normalize_text(v) for kw in keywords)
            )

        if score > best_score:
            best_score = score
            best_sheet = sn

    return best_sheet


def _find_header_row(filepath, keywords, sheet_name=0, max_scan=40):
    """
    Escanea las primeras filas del Excel y devuelve la posición
    de la cabecera basándose en coincidencias con *keywords*.
    Incorpora heurística de desempate: fila con más celdas no vacías
    y mayor ratio texto/numérico.
    """
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name,
                           header=None, nrows=max_scan)
    best_row = None
    best_score = 0

    for idx in range(len(df_raw)):
        row_vals = df_raw.iloc[idx].tolist()
        # Nº de celdas que coinciden con keywords
        kw_score = sum(
            1 for v in row_vals
            if _normalize_text(v)
            and any(kw in _normalize_text(v) for kw in keywords)
        )
        if kw_score >= 2 and kw_score > best_score:
            best_score = kw_score
            best_row = idx

    if best_row is not None:
        return best_row

    # ── Fallback: buscar primera fila donde la mayoría de celdas son texto
    #    (no números ni vacías) — probablemente la cabecera
    for idx in range(len(df_raw)):
        row_vals = df_raw.iloc[idx].tolist()
        non_empty = [v for v in row_vals if pd.notna(v) and str(v).strip()]
        if len(non_empty) < 2:
            continue
        text_count = sum(
            1 for v in non_empty
            if not isinstance(v, (int, float))
            and pd.to_numeric(str(v).strip(), errors='coerce') is np.nan
        )
        if text_count >= len(non_empty) * 0.5:
            return idx

    return 0  # último recurso


def _validate_layout(filepath, header_row, col_map, sheet_name=0,
                     require_date='Fecha', require_numeric=None):
    """
    Valida que las columnas detectadas realmente contienen datos del tipo
    esperado (fechas parseables, números). Devuelve (ok: bool, info: str).
    """
    try:
        df_data = pd.read_excel(
            filepath, sheet_name=sheet_name, header=None,
            skiprows=header_row + 1, nrows=50
        )
    except Exception as e:
        return False, f"Error leyendo datos: {e}"

    issues = []

    # Verificar fecha
    if require_date and require_date in col_map:
        date_col = df_data.iloc[:, col_map[require_date]]
        if not _is_date_col(date_col, 0.2):
            issues.append(
                f"Columna '{require_date}' (pos {col_map[require_date]}) "
                f"no contiene fechas reconocibles"
            )

    # Verificar numérica(s)
    if require_numeric:
        for col_name in require_numeric:
            if col_name in col_map:
                num_col = df_data.iloc[:, col_map[col_name]]
                numeric = _to_numeric_robust(num_col)
                if numeric.notna().sum() / max(len(num_col), 1) < 0.2:
                    issues.append(
                        f"Columna '{col_name}' (pos {col_map[col_name]}) "
                        f"no contiene números reconocibles"
                    )

    if issues:
        return False, '; '.join(issues)
    return True, 'OK'


# ── Funciones de detección principales ────────────────────────────────────────

def detect_banco_layout(filepath):
    """
    Detecta automáticamente la disposición del extracto bancario.

    Devuelve dict:
        header_row   – índice 0-based de la fila de cabecera
        columns      – {nombre_semántico: índice_columna}
        header_names – lista con los nombres originales de las cabeceras
        sheet_name   – nombre de la hoja usada
    """
    sheet = _get_best_sheet(filepath, _BANCO_HEADER_KW)
    header_row = _find_header_row(filepath, _BANCO_HEADER_KW, sheet_name=sheet)

    df_hdr = pd.read_excel(filepath, sheet_name=sheet, header=None,
                           skiprows=header_row, nrows=1)
    header_vals = df_hdr.iloc[0].tolist()
    normalized = [_normalize_text(v) for v in header_vals]

    mapping = {}
    used = set()

    def _first(keywords, exclude=None):
        for i, norm in enumerate(normalized):
            if i in used or not norm:
                continue
            if exclude and any(ex in norm for ex in exclude):
                continue
            if any(kw in norm for kw in keywords):
                return i
        return None

    def _assign(name, idx):
        if idx is not None:
            mapping[name] = idx
            used.add(idx)

    # 1. Fecha
    _assign('Fecha', _first(_KW_FECHA))

    # 2. Importe único (excluir debe/haber/saldo)
    _assign('Importe', _first(_KW_IMPORTE, exclude=_KW_IMPORTE_EXCL))

    # 3. Si no hay importe único, buscar Debe/Haber para banco
    if 'Importe' not in mapping:
        _assign('Banco_Debe', _first(_KW_DEBE))
        _assign('Banco_Haber', _first(_KW_HABER))

    # 4. Movimiento / Descripción
    _assign('Movimiento', _first(_KW_MOVIMIENTO))

    # 5. Más datos
    _assign('Mas_datos', _first(_KW_MAS_DATOS))

    # ── Fallback por inspección de datos ─────────────────────────────────
    if 'Fecha' not in mapping or (
        'Importe' not in mapping
        and 'Banco_Debe' not in mapping
    ):
        try:
            df_data = pd.read_excel(
                filepath, sheet_name=sheet, header=None,
                skiprows=header_row + 1, nrows=100
            )
        except Exception:
            df_data = pd.DataFrame()

        if len(df_data) > 0:
            # Buscar columna de fecha por contenido
            if 'Fecha' not in mapping:
                for i in range(df_data.shape[1]):
                    if i in used:
                        continue
                    if _is_date_col(df_data.iloc[:, i], 0.3):
                        mapping['Fecha'] = i
                        used.add(i)
                        break

            # Buscar columna de importe por contenido numérico
            if 'Importe' not in mapping and 'Banco_Debe' not in mapping:
                numeric_cols = []
                for i in range(df_data.shape[1]):
                    if i in used:
                        continue
                    vals = _to_numeric_robust(df_data.iloc[:, i])
                    pct = vals.notna().sum() / max(len(vals), 1)
                    if pct > 0.3:
                        has_neg = (vals < 0).any()
                        numeric_cols.append((i, pct, has_neg))

                if numeric_cols:
                    # Preferir columna con valores negativos (típico de importes)
                    with_neg = [c for c in numeric_cols if c[2]]
                    if with_neg:
                        mapping['Importe'] = with_neg[0][0]
                    else:
                        # Última columna numérica (suele ser el importe)
                        mapping['Importe'] = numeric_cols[-1][0]
                    used.add(mapping['Importe'])

    # Fallback: primera columna texto no mapeada → Movimiento
    if 'Movimiento' not in mapping:
        excl_kw = _KW_DEBE + _KW_HABER + _KW_SALDO + _KW_IMPORTE
        for i, norm in enumerate(normalized):
            if i in used or not norm:
                continue
            if not any(kw in norm for kw in excl_kw):
                mapping['Movimiento'] = i
                used.add(i)
                break

    # Fallback: segunda columna texto no mapeada → Mas_datos
    if 'Mas_datos' not in mapping:
        excl_kw = _KW_DEBE + _KW_HABER + _KW_SALDO + _KW_IMPORTE
        for i, norm in enumerate(normalized):
            if i in used or not norm:
                continue
            if not any(kw in norm for kw in excl_kw):
                mapping['Mas_datos'] = i
                used.add(i)
                break

    # ── Validación ────────────────────────────────────────────────────────
    req_num = ['Importe'] if 'Importe' in mapping else ['Banco_Debe']
    ok, info = _validate_layout(
        filepath, header_row, mapping, sheet_name=sheet,
        require_date='Fecha', require_numeric=req_num
    )
    if not ok:
        print(f"    ⚠ Validación layout banco: {info}")

    return {
        'header_row': header_row,
        'columns': mapping,
        'header_names': header_vals,
        'sheet_name': sheet,
    }


def detect_mayor_layout(filepath):
    """
    Detecta automáticamente la disposición del libro mayor.

    Columnas semánticas:
        Fecha, Importe_neto (obligatorias),
        Documento, Concepto, Contrapartida, Debe, Haber, Saldo,
        Importe_red, Marca (opcionales).
    Si no existe Importe_neto se calculará a partir de Debe − Haber.
    """
    sheet = _get_best_sheet(filepath, _MAYOR_HEADER_KW)
    header_row = _find_header_row(filepath, _MAYOR_HEADER_KW, sheet_name=sheet)

    df_hdr = pd.read_excel(filepath, sheet_name=sheet, header=None,
                           skiprows=header_row, nrows=1)
    header_vals = df_hdr.iloc[0].tolist()
    normalized = [_normalize_text(v) for v in header_vals]

    mapping = {}
    used = set()

    def _first(keywords, exclude=None):
        for i, norm in enumerate(normalized):
            if i in used or not norm:
                continue
            if exclude and any(ex in norm for ex in exclude):
                continue
            if any(kw in norm for kw in keywords):
                return i
        return None

    def _assign(name, idx):
        if idx is not None:
            mapping[name] = idx
            used.add(idx)

    # Orden: más específico primero para resolver ambigüedades
    _assign('Fecha',  _first(_KW_FECHA))
    _assign('Debe',   _first(_KW_DEBE))
    _assign('Haber',  _first(_KW_HABER))
    _assign('Saldo',  _first(_KW_SALDO))

    # Importe_neto: keyword específico, luego genérico
    idx_neto = _first(_KW_NETO)
    if idx_neto is None:
        idx_neto = _first(_KW_IMPORTE, exclude=_KW_IMPORTE_EXCL)
    _assign('Importe_neto', idx_neto)

    _assign('Importe_red', _first(_KW_RED))
    _assign('Documento',   _first(_KW_DOCUMENTO))
    _assign('Concepto',    _first(_KW_CONCEPTO))
    _assign('Contrapartida', _first(
        _KW_CONTRAPARTIDA,
        exclude=['concepto', 'descripcion']
    ))
    _assign('Marca', _first(_KW_MARCA))

    # ── Fallback por inspección de datos ─────────────────────────────────
    if 'Fecha' not in mapping or (
        'Importe_neto' not in mapping
        and 'Debe' not in mapping
    ):
        try:
            df_data = pd.read_excel(
                filepath, sheet_name=sheet, header=None,
                skiprows=header_row + 1, nrows=100
            )
        except Exception:
            df_data = pd.DataFrame()

        if len(df_data) > 0:
            if 'Fecha' not in mapping:
                for i in range(df_data.shape[1]):
                    if i in used:
                        continue
                    if _is_date_col(df_data.iloc[:, i], 0.3):
                        mapping['Fecha'] = i
                        used.add(i)
                        break

            if 'Importe_neto' not in mapping and 'Debe' not in mapping:
                numeric_cols = []
                for i in range(df_data.shape[1]):
                    if i in used:
                        continue
                    vals = _to_numeric_robust(df_data.iloc[:, i])
                    pct = vals.notna().sum() / max(len(vals), 1)
                    if pct > 0.3:
                        has_neg = (vals < 0).any()
                        numeric_cols.append((i, pct, has_neg))

                if numeric_cols:
                    # Columna con negativos → importe neto
                    with_neg = [c for c in numeric_cols if c[2]]
                    if with_neg:
                        mapping['Importe_neto'] = with_neg[0][0]
                        used.add(mapping['Importe_neto'])

    # ── Validación ────────────────────────────────────────────────────────
    req_num = []
    if 'Importe_neto' in mapping:
        req_num.append('Importe_neto')
    elif 'Debe' in mapping:
        req_num.append('Debe')
    ok, info = _validate_layout(
        filepath, header_row, mapping, sheet_name=sheet,
        require_date='Fecha', require_numeric=req_num
    )
    if not ok:
        print(f"    ⚠ Validación layout mayor: {info}")

    return {
        'header_row': header_row,
        'columns': mapping,
        'header_names': header_vals,
        'sheet_name': sheet,
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 1: CARGA Y LIMPIEZA DE DATOS
# ═══════════════════════════════════════════════════════════════════════════════

def find_input_files(directory):
    """Busca automáticamente los archivos de banco y mayor en el directorio."""
    directory = Path(directory)
    banco_file = None
    mayor_file = None

    for f in sorted(directory.iterdir()):
        if f.is_file() and f.suffix.lower() in ['.xls', '.xlsx']:
            name_upper = f.name.upper()
            if 'BANCO' in name_upper:
                banco_file = f
            elif 'MAYOR' in name_upper:
                mayor_file = f

    return banco_file, mayor_file


def load_banco(filepath):
    """Carga y limpia el extracto bancario (detección automática de layout)."""
    print(f"  📄 Cargando extracto bancario: {Path(filepath).name}")

    # ── Detección de layout ─────────────────────────────────────────────────────
    layout = detect_banco_layout(filepath)
    hr = layout['header_row']
    col_map = layout['columns']
    sheet = layout['sheet_name']

    info = ', '.join(
        f'{name}←\"{layout["header_names"][idx]}\"'
        for name, idx in col_map.items()
    )
    print(f"    → Cabecera en fila {hr + 1} | {info}")

    if 'Fecha' not in col_map:
        raise ValueError(
            "No se encontró columna de FECHA en el extracto bancario. "
            f"Cabeceras detectadas: {layout['header_names']}")

    has_single_importe = 'Importe' in col_map
    has_debe_haber = 'Banco_Debe' in col_map or 'Banco_Haber' in col_map
    if not has_single_importe and not has_debe_haber:
        raise ValueError(
            "No se encontró columna de IMPORTE (ni Debe/Haber) en el extracto "
            f"bancario. Cabeceras detectadas: {layout['header_names']}")

    if has_debe_haber and not has_single_importe:
        print("    ⚠ Sin columna Importe única → se calculará como Debe − Haber")

    # ── Lectura de datos ─────────────────────────────────────────────────────
    df_full = pd.read_excel(filepath, sheet_name=sheet, header=None)
    df_data = df_full.iloc[hr + 1:].copy()
    df_data['_orig_row'] = df_data.index

    # ── DataFrame estandarizado ───────────────────────────────────────────────
    df = pd.DataFrame()
    df['Fecha'] = _parse_dates(df_data.iloc[:, col_map['Fecha']])

    if has_single_importe:
        df['Importe'] = _to_numeric_robust(
            df_data.iloc[:, col_map['Importe']]
        ).round(2)
    else:
        debe = _to_numeric_robust(
            df_data.iloc[:, col_map['Banco_Debe']]
        ).fillna(0) if 'Banco_Debe' in col_map else pd.Series(0, index=df_data.index)
        haber = _to_numeric_robust(
            df_data.iloc[:, col_map['Banco_Haber']]
        ).fillna(0) if 'Banco_Haber' in col_map else pd.Series(0, index=df_data.index)
        df['Importe'] = (debe - haber).round(2)

    if 'Movimiento' in col_map:
        df['Movimiento'] = (
            df_data.iloc[:, col_map['Movimiento']]
            .fillna('').astype(str).str.strip()
        )
    else:
        df['Movimiento'] = ''

    if 'Mas_datos' in col_map:
        df['Mas_datos'] = (
            df_data.iloc[:, col_map['Mas_datos']]
            .fillna('').astype(str).str.strip()
        )
    else:
        df['Mas_datos'] = ''

    df['_orig_row'] = df_data['_orig_row'].values

    # Eliminar filas sin fecha o sin importe
    df = df.dropna(subset=['Fecha', 'Importe']).reset_index(drop=True)
    # Eliminar filas con importe = 0 (a veces son líneas resumen)
    df = df[df['Importe'] != 0].reset_index(drop=True)

    # Trazabilidad
    df['banco_idx'] = range(len(df))
    df['fila_excel'] = df['_orig_row'].astype(int) + 1
    df = df.drop(columns=['_orig_row'])

    if len(df) == 0:
        raise ValueError(
            "No se cargó ningún movimiento bancario. Verifica que el archivo "
            "tiene datos con fechas e importes válidos."
        )

    print(f"    ✓ {len(df)} movimientos cargados "
          f"({df['Fecha'].min().strftime('%d/%m/%Y')} - "
          f"{df['Fecha'].max().strftime('%d/%m/%Y')})")

    return df


def load_mayor(filepath):
    """Carga y limpia el libro mayor (detección automática de layout)."""
    print(f"  📄 Cargando libro mayor: {Path(filepath).name}")

    # ── Detección de layout ─────────────────────────────────────────────────────
    layout = detect_mayor_layout(filepath)
    hr = layout['header_row']
    col_map = layout['columns']
    sheet = layout['sheet_name']

    info = ', '.join(
        f'{name}←\"{layout["header_names"][idx]}\"'
        for name, idx in col_map.items()
    )
    print(f"    → Cabecera en fila {hr + 1} | {info}")

    if 'Fecha' not in col_map:
        raise ValueError(
            "No se encontró columna de FECHA en el libro mayor. "
            f"Cabeceras detectadas: {layout['header_names']}")

    needs_compute_neto = False
    if 'Importe_neto' not in col_map:
        if 'Debe' in col_map and 'Haber' in col_map:
            needs_compute_neto = True
            print("    ⚠ Sin columna Importe_neto → se calculará como Debe − Haber")
        elif 'Debe' in col_map or 'Haber' in col_map:
            needs_compute_neto = True
            print("    ⚠ Sin columna Importe_neto → se calculará desde Debe/Haber")
        else:
            raise ValueError(
                "No se encontró columna de IMPORTE NETO ni columnas Debe/Haber "
                f"en el libro mayor. Cabeceras: {layout['header_names']}")

    # ── Lectura de datos ─────────────────────────────────────────────────────
    df_full = pd.read_excel(filepath, sheet_name=sheet, header=None)
    df_data = df_full.iloc[hr + 1:].copy()
    df_data['_orig_row'] = df_data.index

    # ── DataFrame estandarizado ───────────────────────────────────────────────
    df = pd.DataFrame()

    # Fecha
    df['Fecha'] = _parse_dates(df_data.iloc[:, col_map['Fecha']])

    # Columnas numéricas
    for col_name in ['Debe', 'Haber', 'Saldo', 'Importe_neto', 'Importe_red']:
        if col_name in col_map:
            df[col_name] = _to_numeric_robust(
                df_data.iloc[:, col_map[col_name]]
            )
        else:
            df[col_name] = np.nan

    # Calcular Importe_neto si no existía
    if needs_compute_neto:
        df['Importe_neto'] = (
            df['Debe'].fillna(0) - df['Haber'].fillna(0)
        ).round(2)
    else:
        df['Importe_neto'] = df['Importe_neto'].round(2)

    # Si no hay Importe_red, copiar de Importe_neto
    if df['Importe_red'].isna().all():
        df['Importe_red'] = df['Importe_neto']

    # Columnas de texto
    for col_name in ['Concepto', 'Contrapartida']:
        if col_name in col_map:
            df[col_name] = (
                df_data.iloc[:, col_map[col_name]]
                .fillna('').astype(str).str.strip()
            )
        else:
            df[col_name] = ''

    # Documento (puede ser numérico o texto)
    if 'Documento' in col_map:
        raw_doc = df_data.iloc[:, col_map['Documento']]
        df['Documento_str'] = raw_doc.apply(
            lambda x: '' if pd.isna(x)
            else (f"{x:.10g}" if isinstance(x, (int, float)) else str(x).strip())
        )
    else:
        df['Documento_str'] = ''

    df['_orig_row'] = df_data['_orig_row'].values

    # Eliminar filas sin fecha o sin importe
    df = df.dropna(subset=['Fecha', 'Importe_neto']).reset_index(drop=True)

    # Trazabilidad
    df['mayor_idx'] = range(len(df))
    df['fila_excel'] = df['_orig_row'].astype(int) + 1
    df = df.drop(columns=['_orig_row'])

    # Limpiar texto
    df['Concepto'] = df['Concepto'].fillna('').str.strip()
    df['Contrapartida'] = df['Contrapartida'].fillna('').str.strip()

    if len(df) == 0:
        raise ValueError(
            "No se cargó ningún apunte del mayor. Verifica que el archivo "
            "tiene datos con fechas e importes válidos."
        )

    print(f"    ✓ {len(df)} apuntes cargados "
          f"({df['Fecha'].min().strftime('%d/%m/%Y')} - "
          f"{df['Fecha'].max().strftime('%d/%m/%Y')})")

    # Saldo de apertura
    if 'Saldo' in col_map and not df['Saldo'].isna().all():
        primer_saldo = df.iloc[0]['Saldo']
        primer_importe = df.iloc[0]['Importe_neto']
        saldo_apertura = round(primer_saldo - primer_importe, 2)
        print(f"    ✓ Saldo apertura 572.1: {saldo_apertura:,.2f} €")
    else:
        saldo_apertura = 0.0
        print("    ⚠ No se detectó columna de saldo → saldo apertura = 0.00 €")

    return df, saldo_apertura


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 2: MOTOR DE MATCHING (MULTI-PASS, CONSERVADOR)
# ═══════════════════════════════════════════════════════════════════════════════

class MatchingEngine:
    """
    Motor de conciliación multi-pass con política ultra-conservadora.

    Prioridades:
    1. NUNCA generar falsos positivos (marcar como conciliado algo que no lo es)
    2. Es aceptable generar falsos negativos (no reconocer un match real)

    Pases (de mayor a menor confianza):
    - Pass 1: Importe exacto + misma fecha → CONFIANZA ALTA
    - Pass 2: Importe exacto + fecha ±3 días (par único) → CONFIANZA ALTA
    - Pass 3: Importe exacto + fecha ±15 días (par único) → CONFIANZA MEDIA
    - Pass 4: Importe único global, cualquier fecha → CONFIANZA MEDIA
    - Pass 5: Importes duplicados, emparejamiento por proximidad → CONFIANZA MEDIA
    - Pass 6: Detección de splits 1:N (suma de partes) → CONFIANZA BAJA
    - Pass 7: Todo lo restante → PENDIENTE DE REVISIÓN MANUAL
    """

    def __init__(self, banco_df, mayor_df):
        self.banco = banco_df.copy()
        self.mayor = mayor_df.copy()
        self.matched_banco = set()
        self.matched_mayor = set()
        self.matches = []  # Lista de emparejamientos
        self.splits = []   # Lista de agrupaciones 1:N

    def _unmatched_banco(self):
        """Devuelve los índices de banco aún sin conciliar."""
        return set(self.banco['banco_idx']) - self.matched_banco

    def _unmatched_mayor(self):
        """Devuelve los índices de mayor aún sin conciliar."""
        return set(self.mayor['mayor_idx']) - self.matched_mayor

    def _register_match(self, banco_idx, mayor_idx, confianza, pass_num, detalle=""):
        """Registra un emparejamiento 1:1."""
        self.matched_banco.add(banco_idx)
        self.matched_mayor.add(mayor_idx)
        self.matches.append({
            'banco_idx': banco_idx,
            'mayor_idx': mayor_idx,
            'confianza': confianza,
            'pass': pass_num,
            'detalle': detalle,
        })

    def _register_split(self, banco_idx, mayor_indices, confianza, detalle=""):
        """Registra un emparejamiento 1:N (split)."""
        self.matched_banco.add(banco_idx)
        for midx in mayor_indices:
            self.matched_mayor.add(midx)
        self.splits.append({
            'banco_idx': banco_idx,
            'mayor_indices': list(mayor_indices),
            'confianza': confianza,
            'detalle': detalle,
        })

    def run(self):
        """Ejecuta todos los pases de conciliación."""
        print("\n" + "═" * 70)
        print("  EJECUTANDO MOTOR DE CONCILIACIÓN")
        print("═" * 70)

        self._pass1_exact_amount_exact_date()
        self._pass2_exact_amount_close_date()
        self._pass3_exact_amount_wider_window()
        self._pass4_unique_amount_any_date()
        self._pass5_duplicate_amounts_date_proximity()
        self._pass6_subset_sum_splits()

        self._print_summary()

    # ─── PASS 1: Importe exacto + fecha exacta ────────────────────────────
    def _pass1_exact_amount_exact_date(self):
        """
        Busca pares únicos con el mismo importe Y la misma fecha.
        Solo empareja si hay exactamente UN candidato en cada lado.
        """
        ub = self._unmatched_banco()
        um = self._unmatched_mayor()
        count = 0

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)]
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)]

        # Agrupar por (fecha, importe)
        banco_groups = banco_unm.groupby(['Fecha', 'Importe'])['banco_idx'].apply(list)
        mayor_groups = mayor_unm.groupby(['Fecha', 'Importe_neto'])['mayor_idx'].apply(list)

        for (fecha, importe), b_indices in banco_groups.items():
            m_key = (fecha, importe)
            if m_key in mayor_groups.index:
                m_indices = mayor_groups[m_key]
                # Solo emparejar si hay el mismo número en ambos lados
                # Y emparejar en orden (no hay forma de distinguir, así que 1:1 secuencial)
                n_match = min(len(b_indices), len(m_indices))
                if len(b_indices) == len(m_indices):
                    # Coincidencia perfecta: mismo nº en ambos lados
                    for bi, mi in zip(b_indices, m_indices):
                        self._register_match(
                            bi, mi, 'ALTA', 1,
                            'Importe exacto + fecha exacta'
                        )
                        count += 1
                elif len(b_indices) == 1 and len(m_indices) == 1:
                    # Par único
                    self._register_match(
                        b_indices[0], m_indices[0], 'ALTA', 1,
                        'Importe exacto + fecha exacta (par único)'
                    )
                    count += 1
                # Si hay diferente cantidad > 1, NO emparejar (ambiguo)

        print(f"  Pass 1 │ Importe + fecha exacta     │ {count:>4} conciliados")

    # ─── PASS 2: Importe exacto + fecha ±3 días ───────────────────────────
    def _pass2_exact_amount_close_date(self):
        """
        Para los no conciliados, busca por importe exacto en ventana ±3 días.
        Solo si hay un único candidato en la ventana.
        """
        window = timedelta(days=CONFIG['ventana_dias_pass2'])
        count = 0

        ub = self._unmatched_banco()
        um = self._unmatched_mayor()

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)].copy()
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)].copy()

        # Indexar mayor por importe para búsqueda rápida
        mayor_by_importe = mayor_unm.groupby('Importe_neto')

        for _, brow in banco_unm.iterrows():
            if brow['banco_idx'] in self.matched_banco:
                continue

            importe = brow['Importe']
            fecha = brow['Fecha']

            if importe not in mayor_by_importe.groups:
                continue

            candidates = mayor_by_importe.get_group(importe)
            candidates = candidates[
                ~candidates['mayor_idx'].isin(self.matched_mayor)
            ]

            # Filtrar por ventana de fecha
            in_window = candidates[
                (candidates['Fecha'] >= fecha - window) &
                (candidates['Fecha'] <= fecha + window)
            ]

            if len(in_window) == 1:
                # También verificar que este banco es el único candidato
                # para ese mayor en la ventana
                m_row = in_window.iloc[0]
                m_fecha = m_row['Fecha']
                banco_candidates = banco_unm[
                    (banco_unm['Importe'] == importe) &
                    (~banco_unm['banco_idx'].isin(self.matched_banco)) &
                    (banco_unm['Fecha'] >= m_fecha - window) &
                    (banco_unm['Fecha'] <= m_fecha + window)
                ]
                if len(banco_candidates) == 1:
                    self._register_match(
                        brow['banco_idx'], m_row['mayor_idx'], 'ALTA', 2,
                        f'Importe exacto + fecha ±{CONFIG["ventana_dias_pass2"]}d'
                    )
                    count += 1

        print(f"  Pass 2 │ Importe + fecha ±3 días    │ {count:>4} conciliados")

    # ─── PASS 3: Importe exacto + fecha ±15 días ──────────────────────────
    def _pass3_exact_amount_wider_window(self):
        """
        Ventana más amplia (±15 días). Solo pares únicos en la ventana.
        Confianza MEDIA por el mayor desfase temporal.
        """
        window = timedelta(days=CONFIG['ventana_dias_pass3'])
        count = 0

        ub = self._unmatched_banco()
        um = self._unmatched_mayor()

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)].copy()
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)].copy()

        mayor_by_importe = mayor_unm.groupby('Importe_neto')

        for _, brow in banco_unm.iterrows():
            if brow['banco_idx'] in self.matched_banco:
                continue

            importe = brow['Importe']
            fecha = brow['Fecha']

            if importe not in mayor_by_importe.groups:
                continue

            candidates = mayor_by_importe.get_group(importe)
            candidates = candidates[
                ~candidates['mayor_idx'].isin(self.matched_mayor)
            ]

            in_window = candidates[
                (candidates['Fecha'] >= fecha - window) &
                (candidates['Fecha'] <= fecha + window)
            ]

            if len(in_window) == 1:
                m_row = in_window.iloc[0]
                m_fecha = m_row['Fecha']
                banco_candidates = banco_unm[
                    (banco_unm['Importe'] == importe) &
                    (~banco_unm['banco_idx'].isin(self.matched_banco)) &
                    (banco_unm['Fecha'] >= m_fecha - window) &
                    (banco_unm['Fecha'] <= m_fecha + window)
                ]
                if len(banco_candidates) == 1:
                    dias_diff = abs((fecha - m_fecha).days)
                    self._register_match(
                        brow['banco_idx'], m_row['mayor_idx'], 'MEDIA', 3,
                        f'Importe exacto + fecha ±{CONFIG["ventana_dias_pass3"]}d '
                        f'(desfase: {dias_diff}d)'
                    )
                    count += 1

        print(f"  Pass 3 │ Importe + fecha ±15 días   │ {count:>4} conciliados")

    # ─── PASS 4: Importe globalmente único ─────────────────────────────────
    def _pass4_unique_amount_any_date(self):
        """
        Si un importe aparece exactamente UNA vez en banco sin conciliar
        y UNA vez en mayor sin conciliar → emparejar.
        Sin restricción de fecha (cubre desfases largos).
        """
        window = timedelta(days=CONFIG['ventana_dias_pass4'])
        count = 0

        ub = self._unmatched_banco()
        um = self._unmatched_mayor()

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)]
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)]

        # Contar ocurrencias de cada importe
        banco_counts = banco_unm['Importe'].value_counts()
        mayor_counts = mayor_unm['Importe_neto'].value_counts()

        # Importes que aparecen exactamente 1 vez en cada lado
        banco_unique = set(banco_counts[banco_counts == 1].index)
        mayor_unique = set(mayor_counts[mayor_counts == 1].index)
        common_unique = banco_unique & mayor_unique

        for importe in common_unique:
            b_row = banco_unm[banco_unm['Importe'] == importe].iloc[0]
            m_row = mayor_unm[mayor_unm['Importe_neto'] == importe].iloc[0]

            if b_row['banco_idx'] in self.matched_banco:
                continue
            if m_row['mayor_idx'] in self.matched_mayor:
                continue

            # Verificar que la distancia temporal sea razonable
            dias_diff = abs((b_row['Fecha'] - m_row['Fecha']).days)
            if dias_diff <= CONFIG['ventana_dias_pass4']:
                self._register_match(
                    b_row['banco_idx'], m_row['mayor_idx'], 'MEDIA', 4,
                    f'Importe único global (desfase: {dias_diff}d)'
                )
                count += 1

        print(f"  Pass 4 │ Importe único global       │ {count:>4} conciliados")

    # ─── PASS 5: Importes duplicados + proximidad temporal ─────────────────
    def _pass5_duplicate_amounts_date_proximity(self):
        """
        Para importes que aparecen múltiples veces, empareja por
        proximidad de fecha usando un algoritmo greedy conservador.
        Solo empareja si la distancia ≤ ventana configurada.
        """
        max_days = CONFIG['ventana_dias_duplicados']
        count = 0

        ub = self._unmatched_banco()
        um = self._unmatched_mayor()

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)].copy()
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)].copy()

        # Encontrar importes que aparecen en ambos lados
        banco_importes = set(banco_unm['Importe'].unique())
        mayor_importes = set(mayor_unm['Importe_neto'].unique())
        common_importes = banco_importes & mayor_importes

        for importe in common_importes:
            b_entries = banco_unm[
                (banco_unm['Importe'] == importe) &
                (~banco_unm['banco_idx'].isin(self.matched_banco))
            ].sort_values('Fecha')

            m_entries = mayor_unm[
                (mayor_unm['Importe_neto'] == importe) &
                (~mayor_unm['mayor_idx'].isin(self.matched_mayor))
            ].sort_values('Fecha')

            if len(b_entries) == 0 or len(m_entries) == 0:
                continue

            # Greedy matching por proximidad temporal
            used_mayor = set()
            pairs = []

            for _, brow in b_entries.iterrows():
                best_midx = None
                best_dist = float('inf')

                for _, mrow in m_entries.iterrows():
                    if mrow['mayor_idx'] in used_mayor:
                        continue
                    if mrow['mayor_idx'] in self.matched_mayor:
                        continue
                    dist = abs((brow['Fecha'] - mrow['Fecha']).days)
                    if dist < best_dist:
                        best_dist = dist
                        best_midx = mrow['mayor_idx']

                if best_midx is not None and best_dist <= max_days:
                    pairs.append((brow['banco_idx'], best_midx, best_dist))
                    used_mayor.add(best_midx)

            # Verificar que no hay conflictos en el matching
            # (cada par debe ser mutuamente el mejor)
            for bidx, midx, dist in pairs:
                if bidx not in self.matched_banco and midx not in self.matched_mayor:
                    self._register_match(
                        bidx, midx, 'MEDIA', 5,
                        f'Duplicado emparejado por proximidad ({dist}d)'
                    )
                    count += 1

        print(f"  Pass 5 │ Duplicados + proximidad    │ {count:>4} conciliados")

    # ─── PASS 6: Detección de splits (1:N) ─────────────────────────────────
    def _pass6_subset_sum_splits(self):
        """
        Busca movimientos bancarios cuyo importe coincide con la SUMA
        de 2-N apuntes del mayor (splits tipo Amazon, pagos inmobiliarios, etc.)

        Búsqueda limitada por:
        - Ventana temporal de ±45 días
        - Máximo de 4 partes
        - Solo candidatos del mismo signo
        """
        max_parts = CONFIG['max_split_parts']
        window = timedelta(days=CONFIG['ventana_dias_split'])
        count = 0

        ub = self._unmatched_banco()
        um = self._unmatched_mayor()

        banco_unm = self.banco[self.banco['banco_idx'].isin(ub)].copy()
        mayor_unm = self.mayor[self.mayor['mayor_idx'].isin(um)].copy()

        # Ordenar banco por importe absoluto descendente (priorizar grandes)
        banco_unm = banco_unm.reindex(
            banco_unm['Importe'].abs().sort_values(ascending=False).index
        )

        for _, brow in banco_unm.iterrows():
            if brow['banco_idx'] in self.matched_banco:
                continue

            target = brow['Importe']
            fecha = brow['Fecha']

            # Filtrar candidatos del mayor
            candidates = mayor_unm[
                (~mayor_unm['mayor_idx'].isin(self.matched_mayor)) &
                (mayor_unm['Fecha'] >= fecha - window) &
                (mayor_unm['Fecha'] <= fecha + window)
            ]

            # Solo mismo signo y |valor individual| ≤ |target|
            if target < 0:
                candidates = candidates[
                    (candidates['Importe_neto'] < 0) &
                    (candidates['Importe_neto'] >= target - 0.01)
                ]
            else:
                candidates = candidates[
                    (candidates['Importe_neto'] > 0) &
                    (candidates['Importe_neto'] <= target + 0.01)
                ]

            if len(candidates) < 2:
                continue

            # Limitar candidatos para viabilidad computacional
            if len(candidates) > 60:
                # Priorizar los más cercanos en fecha
                candidates = candidates.copy()
                candidates['_fecha_dist'] = abs(
                    (candidates['Fecha'] - fecha).dt.days
                )
                candidates = candidates.nsmallest(60, '_fecha_dist')

            cand_list = list(zip(
                candidates['mayor_idx'],
                candidates['Importe_neto']
            ))

            # Buscar subsets que sumen al target
            found_subsets = []
            for size in range(2, min(max_parts + 1, len(cand_list) + 1)):
                if found_subsets:
                    break  # Ya encontramos en tamaño menor, no buscar más
                for combo in combinations(cand_list, size):
                    combo_sum = round(sum(amt for _, amt in combo), 2)
                    if combo_sum == target:
                        found_subsets.append(combo)
                        if len(found_subsets) > 5:
                            break  # Demasiadas opciones = ambiguo

            # Filtrar subsets: las partes del mayor deben estar
            # temporalmente agrupadas (max 7 días entre la primera y última)
            valid_subsets = []
            for subset in found_subsets:
                m_indices = [idx for idx, _ in subset]
                m_dates = [self.mayor.loc[self.mayor['mayor_idx'] == idx, 'Fecha'].iloc[0]
                           for idx in m_indices]
                date_spread = (max(m_dates) - min(m_dates)).days
                if date_spread <= 7:
                    valid_subsets.append(subset)

            if len(valid_subsets) == 1:
                # Exactamente un subset válido → registrar como split
                indices = [idx for idx, _ in valid_subsets[0]]
                self._register_split(
                    brow['banco_idx'], indices, 'BAJA',
                    f'Split 1:{len(indices)} detectado '
                    f'({" + ".join(f"{amt:.2f}" for _, amt in valid_subsets[0])} = {target:.2f})'
                )
                count += 1
            elif len(valid_subsets) > 1:
                # Múltiples opciones → NO emparejar, demasiado ambiguo
                pass

        print(f"  Pass 6 │ Splits (1:N)               │ {count:>4} conciliados")

    # ─── RESUMEN ───────────────────────────────────────────────────────────
    def _print_summary(self):
        """Imprime resumen de la conciliación."""
        total_b = len(self.banco)
        total_m = len(self.mayor)
        matched_b = len(self.matched_banco)
        matched_m = len(self.matched_mayor)
        unmatched_b = total_b - matched_b
        unmatched_m = total_m - matched_m

        print("  " + "─" * 68)
        print(f"  TOTAL  │ Banco: {matched_b}/{total_b} conciliados "
              f"│ Mayor: {matched_m}/{total_m} conciliados")
        print(f"         │ Banco pendiente: {unmatched_b:>4}        "
              f"│ Mayor pendiente: {unmatched_m:>4}")
        print("═" * 70)

    def get_results(self):
        """Devuelve los resultados estructurados."""
        return {
            'matches': self.matches,
            'splits': self.splits,
            'matched_banco': self.matched_banco,
            'matched_mayor': self.matched_mayor,
            'banco': self.banco,
            'mayor': self.mayor,
        }


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 3: GENERACIÓN DEL INFORME EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def generate_report(results, saldo_apertura, output_path):
    """Genera el Excel de salida con todas las hojas de detalle."""

    banco = results['banco']
    mayor = results['mayor']
    matches = results['matches']
    splits = results['splits']
    matched_banco = results['matched_banco']
    matched_mayor = results['matched_mayor']

    fecha_inicio = pd.Timestamp(CONFIG['fecha_inicio'])
    fecha_fin = pd.Timestamp(CONFIG['fecha_fin'])

    print(f"\n  📊 Generando informe: {Path(output_path).name}")

    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book

    # ─── Formatos ──────────────────────────────────────────────────────
    fmt_header = workbook.add_format({
        'bold': True, 'bg_color': '#2F5496', 'font_color': 'white',
        'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        'font_size': 10,
    })
    fmt_alta = workbook.add_format({
        'bg_color': '#C6EFCE', 'font_color': '#006100',
        'border': 1, 'font_size': 9,
    })
    fmt_media = workbook.add_format({
        'bg_color': '#FFEB9C', 'font_color': '#9C5700',
        'border': 1, 'font_size': 9,
    })
    fmt_baja = workbook.add_format({
        'bg_color': '#FCE4D6', 'font_color': '#BF4012',
        'border': 1, 'font_size': 9,
    })
    fmt_pendiente = workbook.add_format({
        'bg_color': '#FFC7CE', 'font_color': '#9C0006',
        'border': 1, 'font_size': 9,
    })
    fmt_normal = workbook.add_format({
        'border': 1, 'font_size': 9,
    })
    fmt_money = workbook.add_format({
        'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
    })
    fmt_money_green = workbook.add_format({
        'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
        'bg_color': '#C6EFCE', 'font_color': '#006100',
    })
    fmt_money_yellow = workbook.add_format({
        'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
        'bg_color': '#FFEB9C', 'font_color': '#9C5700',
    })
    fmt_money_red = workbook.add_format({
        'border': 1, 'font_size': 9, 'num_format': '#,##0.00',
        'bg_color': '#FFC7CE', 'font_color': '#9C0006',
    })
    fmt_date = workbook.add_format({
        'border': 1, 'font_size': 9, 'num_format': 'dd/mm/yyyy',
    })
    fmt_title = workbook.add_format({
        'bold': True, 'font_size': 14, 'font_color': '#2F5496',
    })
    fmt_subtitle = workbook.add_format({
        'bold': True, 'font_size': 11, 'font_color': '#2F5496',
    })
    fmt_kpi = workbook.add_format({
        'bold': True, 'font_size': 12, 'num_format': '#,##0.00',
    })
    fmt_kpi_pct = workbook.add_format({
        'bold': True, 'font_size': 12, 'num_format': '0.0%',
    })

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 1: RESUMEN
    # ═══════════════════════════════════════════════════════════════════
    ws = workbook.add_worksheet('RESUMEN')
    writer.sheets['RESUMEN'] = ws
    ws.set_column('A:A', 40)
    ws.set_column('B:B', 20)
    ws.set_column('C:C', 20)
    ws.hide_gridlines(2)

    row = 0
    ws.write(row, 0, 'CONCILIACIÓN BANCARIA AUTOMATIZADA', fmt_title)
    row += 1
    ws.write(row, 0, f'Generado: {datetime.now().strftime("%d/%m/%Y %H:%M")}',
             fmt_normal)
    row += 2

    # Filtrar período principal
    banco_periodo = banco[
        (banco['Fecha'] >= fecha_inicio) & (banco['Fecha'] <= fecha_fin)
    ]
    mayor_periodo = mayor[
        (mayor['Fecha'] >= fecha_inicio) & (mayor['Fecha'] <= fecha_fin)
    ]
    matched_b_periodo = matched_banco & set(banco_periodo['banco_idx'])
    matched_m_periodo = matched_mayor & set(mayor_periodo['mayor_idx'])

    ws.write(row, 0, f'PERÍODO: {CONFIG["fecha_inicio"]} a {CONFIG["fecha_fin"]}',
             fmt_subtitle)
    row += 2

    # Estadísticas banco
    ws.write(row, 0, 'EXTRACTO BANCARIO', fmt_subtitle)
    row += 1
    stats = [
        ('Movimientos totales (período)', len(banco_periodo), None),
        ('Movimientos conciliados', len(matched_b_periodo), None),
        ('Movimientos PENDIENTES', len(banco_periodo) - len(matched_b_periodo), None),
        ('% conciliado', len(matched_b_periodo) / max(len(banco_periodo), 1), 'pct'),
        ('Suma total banco (período)',
         banco_periodo['Importe'].sum(), 'money'),
        ('Suma conciliada',
         banco[banco['banco_idx'].isin(matched_b_periodo)]['Importe'].sum(), 'money'),
        ('Suma PENDIENTE',
         banco[
             (banco['banco_idx'].isin(set(banco_periodo['banco_idx']))) &
             (~banco['banco_idx'].isin(matched_banco))
         ]['Importe'].sum(), 'money'),
    ]
    for label, value, tipo in stats:
        ws.write(row, 0, label, fmt_normal)
        if tipo == 'pct':
            ws.write(row, 1, value, fmt_kpi_pct)
        elif tipo == 'money':
            ws.write(row, 1, value, fmt_kpi)
        else:
            ws.write(row, 1, value, fmt_kpi)
        row += 1

    row += 1
    ws.write(row, 0, 'LIBRO MAYOR 572.1', fmt_subtitle)
    row += 1
    stats_m = [
        ('Apuntes totales (período)', len(mayor_periodo), None),
        ('Apuntes conciliados', len(matched_m_periodo), None),
        ('Apuntes PENDIENTES', len(mayor_periodo) - len(matched_m_periodo), None),
        ('% conciliado', len(matched_m_periodo) / max(len(mayor_periodo), 1), 'pct'),
        ('Suma Importe neto (período)',
         mayor_periodo['Importe_neto'].sum(), 'money'),
    ]
    for label, value, tipo in stats_m:
        ws.write(row, 0, label, fmt_normal)
        if tipo == 'pct':
            ws.write(row, 1, value, fmt_kpi_pct)
        elif tipo == 'money':
            ws.write(row, 1, value, fmt_kpi)
        else:
            ws.write(row, 1, value, fmt_kpi)
        row += 1

    row += 1
    ws.write(row, 0, 'CONCILIACIÓN POR NIVEL DE CONFIANZA', fmt_subtitle)
    row += 1
    for nivel in ['ALTA', 'MEDIA', 'BAJA']:
        n = sum(1 for m in matches if m['confianza'] == nivel)
        ws.write(row, 0, f'  Confianza {nivel}', fmt_normal)
        ws.write(row, 1, n, fmt_kpi)
        row += 1
    n_splits = len(splits)
    ws.write(row, 0, '  Agrupaciones (1:N)', fmt_normal)
    ws.write(row, 1, n_splits, fmt_kpi)
    row += 2

    # Saldo reconciliation
    ws.write(row, 0, 'VERIFICACIÓN DE SALDOS', fmt_subtitle)
    row += 1
    ws.write(row, 0, 'Saldo apertura 572.1', fmt_normal)
    ws.write(row, 1, saldo_apertura, fmt_kpi)
    row += 1

    mayor_2025 = mayor[
        (mayor['Fecha'] >= fecha_inicio) & (mayor['Fecha'] <= fecha_fin)
    ]
    ultimo_idx = mayor_2025['Fecha'].idxmax() if len(mayor_2025) > 0 else None
    if ultimo_idx is not None:
        ultimo_saldo = mayor_2025.loc[ultimo_idx, 'Saldo']
        ws.write(row, 0,
                 f'Último saldo mayor ({mayor_2025.loc[ultimo_idx, "Fecha"].strftime("%d/%m/%Y")})',
                 fmt_normal)
        ws.write(row, 1, ultimo_saldo, fmt_kpi)
    row += 1

    suma_banco_2025 = banco_periodo['Importe'].sum()
    ws.write(row, 0, 'Suma movimientos banco 2025', fmt_normal)
    ws.write(row, 1, suma_banco_2025, fmt_kpi)
    row += 1
    suma_mayor_2025 = mayor_periodo['Importe_neto'].sum()
    ws.write(row, 0, 'Suma movimientos mayor 2025', fmt_normal)
    ws.write(row, 1, suma_mayor_2025, fmt_kpi)
    row += 1
    ws.write(row, 0, 'Diferencia (banco - mayor)', fmt_normal)
    ws.write(row, 1, round(suma_banco_2025 - suma_mayor_2025, 2), fmt_kpi)

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 2: CONCILIADOS (parejas encontradas)
    # ═══════════════════════════════════════════════════════════════════
    ws2 = workbook.add_worksheet('CONCILIADOS')
    writer.sheets['CONCILIADOS'] = ws2

    headers_conc = [
        'Confianza', 'Pass', 'Detalle',
        'B_Fila', 'B_Fecha', 'B_Movimiento', 'B_Mas_datos', 'B_Importe',
        'M_Fila', 'M_Fecha', 'M_Documento', 'M_Concepto', 'M_Contrapartida',
        'M_Importe', 'Diferencia_días'
    ]
    for col, h in enumerate(headers_conc):
        ws2.write(0, col, h, fmt_header)

    ws2.set_column('A:A', 12)
    ws2.set_column('B:B', 6)
    ws2.set_column('C:C', 35)
    ws2.set_column('D:D', 8)
    ws2.set_column('E:E', 12)
    ws2.set_column('F:F', 22)
    ws2.set_column('G:G', 30)
    ws2.set_column('H:H', 14)
    ws2.set_column('I:I', 8)
    ws2.set_column('J:J', 12)
    ws2.set_column('K:K', 12)
    ws2.set_column('L:L', 45)
    ws2.set_column('M:M', 45)
    ws2.set_column('N:N', 14)
    ws2.set_column('O:O', 10)
    ws2.autofilter(0, 0, 0, len(headers_conc) - 1)
    ws2.freeze_panes(1, 0)

    row = 1
    for m in sorted(matches, key=lambda x: (x['pass'], x['banco_idx'])):
        brow = banco[banco['banco_idx'] == m['banco_idx']].iloc[0]
        mrow = mayor[mayor['mayor_idx'] == m['mayor_idx']].iloc[0]

        confianza = m['confianza']
        if confianza == 'ALTA':
            fmt_row = fmt_alta
            fmt_m = fmt_money_green
        elif confianza == 'MEDIA':
            fmt_row = fmt_media
            fmt_m = fmt_money_yellow
        else:
            fmt_row = fmt_baja
            fmt_m = fmt_money

        dias_diff = abs((brow['Fecha'] - mrow['Fecha']).days)

        ws2.write(row, 0, confianza, fmt_row)
        ws2.write(row, 1, m['pass'], fmt_row)
        ws2.write(row, 2, m['detalle'], fmt_row)
        ws2.write(row, 3, int(brow['fila_excel']), fmt_row)
        ws2.write_datetime(row, 4, brow['Fecha'].to_pydatetime(), fmt_date)
        ws2.write(row, 5, brow['Movimiento'], fmt_row)
        ws2.write(row, 6, brow['Mas_datos'], fmt_row)
        ws2.write(row, 7, brow['Importe'], fmt_m)
        ws2.write(row, 8, int(mrow['fila_excel']), fmt_row)
        ws2.write_datetime(row, 9, mrow['Fecha'].to_pydatetime(), fmt_date)
        ws2.write(row, 10, mrow['Documento_str'], fmt_row)
        ws2.write(row, 11, mrow['Concepto'], fmt_row)
        ws2.write(row, 12, mrow['Contrapartida'], fmt_row)
        ws2.write(row, 13, mrow['Importe_neto'], fmt_m)
        ws2.write(row, 14, dias_diff, fmt_row)
        row += 1

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 3: BANCO NO CONCILIADO
    # ═══════════════════════════════════════════════════════════════════
    ws3 = workbook.add_worksheet('BANCO_PENDIENTE')
    writer.sheets['BANCO_PENDIENTE'] = ws3

    headers_bp = ['Fila_Excel', 'Fecha', 'Movimiento', 'Mas_datos', 'Importe',
                  'Periodo']
    for col, h in enumerate(headers_bp):
        ws3.write(0, col, h, fmt_header)

    ws3.set_column('A:A', 10)
    ws3.set_column('B:B', 12)
    ws3.set_column('C:C', 25)
    ws3.set_column('D:D', 35)
    ws3.set_column('E:E', 14)
    ws3.set_column('F:F', 12)
    ws3.autofilter(0, 0, 0, len(headers_bp) - 1)
    ws3.freeze_panes(1, 0)

    banco_pendiente = banco[~banco['banco_idx'].isin(matched_banco)].sort_values('Fecha')
    row = 1
    for _, brow in banco_pendiente.iterrows():
        periodo = '2025' if fecha_inicio <= brow['Fecha'] <= fecha_fin else 'Otro'
        ws3.write(row, 0, int(brow['fila_excel']), fmt_pendiente)
        ws3.write_datetime(row, 1, brow['Fecha'].to_pydatetime(), fmt_date)
        ws3.write(row, 2, brow['Movimiento'], fmt_pendiente)
        ws3.write(row, 3, brow['Mas_datos'], fmt_pendiente)
        ws3.write(row, 4, brow['Importe'], fmt_money_red)
        ws3.write(row, 5, periodo, fmt_pendiente)
        row += 1

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 4: MAYOR NO CONCILIADO
    # ═══════════════════════════════════════════════════════════════════
    ws4 = workbook.add_worksheet('MAYOR_PENDIENTE')
    writer.sheets['MAYOR_PENDIENTE'] = ws4

    headers_mp = ['Fila_Excel', 'Fecha', 'Documento', 'Concepto',
                  'Contrapartida', 'Debe', 'Haber', 'Importe_neto', 'Periodo']
    for col, h in enumerate(headers_mp):
        ws4.write(0, col, h, fmt_header)

    ws4.set_column('A:A', 10)
    ws4.set_column('B:B', 12)
    ws4.set_column('C:C', 10)
    ws4.set_column('D:D', 55)
    ws4.set_column('E:E', 55)
    ws4.set_column('F:F', 14)
    ws4.set_column('G:G', 14)
    ws4.set_column('H:H', 14)
    ws4.set_column('I:I', 12)
    ws4.autofilter(0, 0, 0, len(headers_mp) - 1)
    ws4.freeze_panes(1, 0)

    mayor_pendiente = mayor[~mayor['mayor_idx'].isin(matched_mayor)].sort_values('Fecha')
    row = 1
    for _, mrow in mayor_pendiente.iterrows():
        periodo = '2025' if fecha_inicio <= mrow['Fecha'] <= fecha_fin else 'Otro'
        debe_val = mrow['Debe'] if pd.notna(mrow['Debe']) else ''
        haber_val = mrow['Haber'] if pd.notna(mrow['Haber']) else ''

        ws4.write(row, 0, int(mrow['fila_excel']), fmt_pendiente)
        ws4.write_datetime(row, 1, mrow['Fecha'].to_pydatetime(), fmt_date)
        ws4.write(row, 2, mrow['Documento_str'], fmt_pendiente)
        ws4.write(row, 3, mrow['Concepto'], fmt_pendiente)
        ws4.write(row, 4, mrow['Contrapartida'], fmt_pendiente)
        if debe_val != '':
            ws4.write(row, 5, debe_val, fmt_money_red)
        else:
            ws4.write(row, 5, '', fmt_pendiente)
        if haber_val != '':
            ws4.write(row, 6, haber_val, fmt_money_red)
        else:
            ws4.write(row, 6, '', fmt_pendiente)
        ws4.write(row, 7, mrow['Importe_neto'], fmt_money_red)
        ws4.write(row, 8, periodo, fmt_pendiente)
        row += 1

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 5: AGRUPACIONES / SPLITS (1:N)
    # ═══════════════════════════════════════════════════════════════════
    ws5 = workbook.add_worksheet('AGRUPACIONES_1N')
    writer.sheets['AGRUPACIONES_1N'] = ws5

    headers_split = [
        'Grupo', 'Confianza', 'Detalle',
        'B_Fila', 'B_Fecha', 'B_Movimiento', 'B_Mas_datos', 'B_Importe',
        'M_Fila', 'M_Fecha', 'M_Documento', 'M_Concepto', 'M_Contrapartida',
        'M_Importe'
    ]
    for col, h in enumerate(headers_split):
        ws5.write(0, col, h, fmt_header)

    ws5.set_column('A:A', 8)
    ws5.set_column('B:B', 12)
    ws5.set_column('C:C', 50)
    ws5.set_column('D:D', 8)
    ws5.set_column('E:E', 12)
    ws5.set_column('F:F', 22)
    ws5.set_column('G:G', 30)
    ws5.set_column('H:H', 14)
    ws5.set_column('I:I', 8)
    ws5.set_column('J:J', 12)
    ws5.set_column('K:K', 12)
    ws5.set_column('L:L', 50)
    ws5.set_column('M:M', 50)
    ws5.set_column('N:N', 14)
    ws5.autofilter(0, 0, 0, len(headers_split) - 1)
    ws5.freeze_panes(1, 0)

    row = 1
    for grupo_id, split in enumerate(splits, 1):
        brow = banco[banco['banco_idx'] == split['banco_idx']].iloc[0]

        for i, midx in enumerate(split['mayor_indices']):
            mrow = mayor[mayor['mayor_idx'] == midx].iloc[0]

            ws5.write(row, 0, grupo_id, fmt_baja)
            ws5.write(row, 1, split['confianza'], fmt_baja)
            ws5.write(row, 2, split['detalle'] if i == 0 else '', fmt_baja)
            ws5.write(row, 3, int(brow['fila_excel']) if i == 0 else '', fmt_baja)
            if i == 0:
                ws5.write_datetime(row, 4, brow['Fecha'].to_pydatetime(), fmt_date)
            else:
                ws5.write(row, 4, '', fmt_baja)
            ws5.write(row, 5, brow['Movimiento'] if i == 0 else '', fmt_baja)
            ws5.write(row, 6, brow['Mas_datos'] if i == 0 else '', fmt_baja)
            ws5.write(row, 7, brow['Importe'] if i == 0 else '', fmt_baja)
            ws5.write(row, 8, int(mrow['fila_excel']), fmt_baja)
            ws5.write_datetime(row, 9, mrow['Fecha'].to_pydatetime(), fmt_date)
            ws5.write(row, 10, mrow['Documento_str'], fmt_baja)
            ws5.write(row, 11, mrow['Concepto'], fmt_baja)
            ws5.write(row, 12, mrow['Contrapartida'], fmt_baja)
            ws5.write(row, 13, mrow['Importe_neto'], fmt_money)
            row += 1

    # ═══════════════════════════════════════════════════════════════════
    #  HOJA 6: DETALLE COMPLETO (todos los movimientos con estado)
    # ═══════════════════════════════════════════════════════════════════
    ws6 = workbook.add_worksheet('BANCO_COMPLETO')
    writer.sheets['BANCO_COMPLETO'] = ws6

    headers_full = [
        'Fila_Excel', 'Fecha', 'Movimiento', 'Mas_datos', 'Importe',
        'Estado', 'Confianza', 'Pass', 'Mayor_Documento', 'Mayor_Concepto'
    ]
    for col, h in enumerate(headers_full):
        ws6.write(0, col, h, fmt_header)

    ws6.set_column('A:A', 10)
    ws6.set_column('B:B', 12)
    ws6.set_column('C:C', 22)
    ws6.set_column('D:D', 30)
    ws6.set_column('E:E', 14)
    ws6.set_column('F:F', 16)
    ws6.set_column('G:G', 12)
    ws6.set_column('H:H', 6)
    ws6.set_column('I:I', 12)
    ws6.set_column('J:J', 50)
    ws6.autofilter(0, 0, 0, len(headers_full) - 1)
    ws6.freeze_panes(1, 0)

    # Construir lookup de matches
    match_lookup = {}
    for m in matches:
        match_lookup[m['banco_idx']] = m
    split_lookup = {}
    for s in splits:
        split_lookup[s['banco_idx']] = s

    row = 1
    for _, brow in banco.sort_values('Fecha').iterrows():
        bidx = brow['banco_idx']

        if bidx in match_lookup:
            m = match_lookup[bidx]
            mrow = mayor[mayor['mayor_idx'] == m['mayor_idx']].iloc[0]
            estado = 'CONCILIADO'
            confianza = m['confianza']
            pass_num = m['pass']
            doc = mrow['Documento_str']
            concepto = mrow['Concepto']
            fmt_r = fmt_alta if confianza == 'ALTA' else (
                fmt_media if confianza == 'MEDIA' else fmt_baja)
            fmt_mr = fmt_money_green if confianza == 'ALTA' else fmt_money_yellow
        elif bidx in split_lookup:
            s = split_lookup[bidx]
            estado = 'AGRUPACIÓN'
            confianza = s['confianza']
            pass_num = 6
            docs = []
            for midx in s['mayor_indices']:
                mrow_s = mayor[mayor['mayor_idx'] == midx].iloc[0]
                docs.append(mrow_s['Documento_str'])
            doc = ' + '.join(docs)
            concepto = s['detalle']
            fmt_r = fmt_baja
            fmt_mr = fmt_money
        else:
            estado = 'PENDIENTE'
            confianza = ''
            pass_num = ''
            doc = ''
            concepto = ''
            fmt_r = fmt_pendiente
            fmt_mr = fmt_money_red

        ws6.write(row, 0, int(brow['fila_excel']), fmt_r)
        ws6.write_datetime(row, 1, brow['Fecha'].to_pydatetime(), fmt_date)
        ws6.write(row, 2, brow['Movimiento'], fmt_r)
        ws6.write(row, 3, brow['Mas_datos'], fmt_r)
        ws6.write(row, 4, brow['Importe'], fmt_mr)
        ws6.write(row, 5, estado, fmt_r)
        ws6.write(row, 6, confianza, fmt_r)
        ws6.write(row, 7, pass_num, fmt_r)
        ws6.write(row, 8, doc, fmt_r)
        ws6.write(row, 9, concepto, fmt_r)
        row += 1

    # Cerrar
    writer.close()
    print(f"    ✓ Informe generado correctamente")


# ═══════════════════════════════════════════════════════════════════════════════
#  SECCIÓN 4: FUNCIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    print()
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║   SISTEMA DE CONCILIACIÓN BANCARIA AUTOMATIZADA                ║")
    print("║   Punteo: Extracto Bancario ↔ Libro Mayor (572.1)             ║")
    print("╚══════════════════════════════════════════════════════════════════╝")
    print()

    # Detectar directorio
    script_dir = Path(__file__).parent
    print(f"  📁 Directorio de trabajo: {script_dir}")

    # Buscar archivos
    banco_file, mayor_file = find_input_files(script_dir)

    if not banco_file:
        print("  ❌ ERROR: No se encontró archivo del BANCO (*.xls/*.xlsx con 'BANCO' en el nombre)")
        sys.exit(1)
    if not mayor_file:
        print("  ❌ ERROR: No se encontró archivo del MAYOR (*.xls/*.xlsx con 'MAYOR' en el nombre)")
        sys.exit(1)

    # Cargar datos
    print(f"\n  {'─' * 60}")
    print("  CARGANDO DATOS")
    print(f"  {'─' * 60}")
    banco_df = load_banco(banco_file)
    mayor_df, saldo_apertura = load_mayor(mayor_file)

    # Ejecutar conciliación
    engine = MatchingEngine(banco_df, mayor_df)
    engine.run()

    results = engine.get_results()

    # Generar informe
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    # Extraer nombre de empresa del fichero
    stem = banco_file.stem.upper()
    for word in ['BANCOS', 'BANCO', '572.1', '572_1']:
        stem = stem.replace(word, '')
    empresa = stem.strip().strip('_').strip('-').strip() or 'EMPRESA'
    output_name = f"CONCILIACION_{empresa}_{CONFIG['fecha_fin'][:4]}_{timestamp}.xlsx"
    output_path = script_dir / output_name

    generate_report(results, saldo_apertura, output_path)

    # ─── Resumen final en consola ──────────────────────────────────────
    fecha_inicio = pd.Timestamp(CONFIG['fecha_inicio'])
    fecha_fin = pd.Timestamp(CONFIG['fecha_fin'])

    banco_periodo = banco_df[
        (banco_df['Fecha'] >= fecha_inicio) & (banco_df['Fecha'] <= fecha_fin)
    ]
    mayor_periodo = mayor_df[
        (mayor_df['Fecha'] >= fecha_inicio) & (mayor_df['Fecha'] <= fecha_fin)
    ]

    matched_b_periodo = results['matched_banco'] & set(banco_periodo['banco_idx'])
    matched_m_periodo = results['matched_mayor'] & set(mayor_periodo['mayor_idx'])

    unmatched_b_periodo = len(banco_periodo) - len(matched_b_periodo)
    unmatched_m_periodo = len(mayor_periodo) - len(matched_m_periodo)

    print(f"\n  {'═' * 60}")
    print(f"  RESULTADO FINAL — Período {CONFIG['fecha_inicio'][:4]}")
    print(f"  {'═' * 60}")
    print(f"  Banco: {len(matched_b_periodo)}/{len(banco_periodo)} conciliados "
          f"({len(matched_b_periodo)/max(len(banco_periodo),1)*100:.1f}%)")
    print(f"  Mayor: {len(matched_m_periodo)}/{len(mayor_periodo)} conciliados "
          f"({len(matched_m_periodo)/max(len(mayor_periodo),1)*100:.1f}%)")
    print(f"  Banco PENDIENTE: {unmatched_b_periodo} movimientos")
    print(f"  Mayor PENDIENTE: {unmatched_m_periodo} apuntes")

    suma_pend_b = banco_df[
        (banco_df['banco_idx'].isin(set(banco_periodo['banco_idx']))) &
        (~banco_df['banco_idx'].isin(results['matched_banco']))
    ]['Importe'].sum()

    suma_pend_m = mayor_df[
        (mayor_df['mayor_idx'].isin(set(mayor_periodo['mayor_idx']))) &
        (~mayor_df['mayor_idx'].isin(results['matched_mayor']))
    ]['Importe_neto'].sum()

    print(f"\n  Suma pendiente banco: {suma_pend_b:>12,.2f} €")
    print(f"  Suma pendiente mayor: {suma_pend_m:>12,.2f} €")
    print(f"  Diferencia pendiente: {suma_pend_b - suma_pend_m:>12,.2f} €")

    print(f"\n  📄 Informe guardado en: {output_path.name}")
    print()


if __name__ == '__main__':
    main()
