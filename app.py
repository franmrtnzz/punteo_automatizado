#!/usr/bin/env python3
"""
Plataforma de conciliación bancaria — Interfaz local.
Puerto fijo: 8502
"""

import streamlit as st
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

from conciliacion import (
    CONFIG, load_banco, load_mayor, MatchingEngine, generate_report
)
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
#  Página
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Conciliación Bancaria",
    page_icon=" ",
    layout="centered",
)

# ─────────────────────────────────────────────────────────────────────────────
#  Estilo
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    .block-container {
        max-width: 780px;
        padding-top: 2.5rem;
        padding-bottom: 2rem;
    }

    /* Cabecera */
    .app-header {
        border-bottom: 2px solid #1a1a2e;
        padding-bottom: 1rem;
        margin-bottom: 1.8rem;
    }
    .app-header h1 {
        font-size: 1.55rem;
        font-weight: 600;
        color: #1a1a2e;
        margin: 0;
        letter-spacing: -0.02em;
    }
    .app-header p {
        font-size: 0.85rem;
        color: #6b7280;
        margin: 0.3rem 0 0 0;
    }

    /* Uploaders */
    div[data-testid="stFileUploader"] {
        border: 1.5px solid #d1d5db;
        border-radius: 6px;
        padding: 0.4rem;
        background: #fafafa;
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #1a1a2e;
    }

    /* Botón primario */
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="stBaseButton-primary"] {
        width: 100%;
        background-color: #1a1a2e;
        color: #ffffff;
        font-weight: 500;
        font-size: 0.95rem;
        padding: 0.65rem 1.2rem;
        border: none;
        border-radius: 5px;
        letter-spacing: 0.01em;
        transition: background-color 0.15s ease;
    }
    .stButton > button[kind="primary"]:hover,
    .stButton > button[data-testid="stBaseButton-primary"]:hover {
        background-color: #2d2d4a;
        color: #ffffff;
    }

    /* Botón descarga */
    .stDownloadButton > button {
        width: 100%;
        background-color: #1a1a2e;
        color: #ffffff;
        font-weight: 500;
        font-size: 0.95rem;
        padding: 0.65rem 1.2rem;
        border: none;
        border-radius: 5px;
    }
    .stDownloadButton > button:hover {
        background-color: #2d2d4a;
        color: #ffffff;
    }

    /* Métricas */
    div[data-testid="stMetric"] {
        background: #f8f9fa;
        border: 1px solid #e5e7eb;
        border-radius: 6px;
        padding: 0.8rem 1rem;
    }
    div[data-testid="stMetric"] label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #6b7280;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 1.35rem;
        font-weight: 600;
        color: #1a1a2e;
    }
    div[data-testid="stMetric"] [data-testid="stMetricDelta"] {
        color: #374151;
    }

    /* Resultados */
    .results-ok {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 6px;
        padding: 0.75rem 1rem;
        color: #166534;
        font-weight: 500;
        font-size: 0.9rem;
        margin-bottom: 1rem;
    }
    .results-summary {
        font-size: 0.85rem;
        color: #4b5563;
        margin: 0.8rem 0 1.2rem 0;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #f9fafb;
        border-right: 1px solid #e5e7eb;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    /* Ocultar menú hamburguesa y footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {
        background: transparent;
    }

    /* Info box */
    div[data-testid="stNotification"] {
        background: #f3f4f6;
        border: 1px solid #d1d5db;
        color: #374151;
        border-radius: 6px;
    }

    /* Divider */
    hr {
        border-color: #e5e7eb;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  Cabecera
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
    <h1>Conciliación Bancaria</h1>
    <p>Extracto bancario y libro mayor 572.1 — Punteo automatizado</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  Sidebar: parámetros
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("#### Parámetros")

    año = st.number_input("Ejercicio", min_value=2020, max_value=2030,
                          value=2025, step=1)

    with st.expander("Ventanas de matching (dias)"):
        v_pass2 = st.number_input("Pass 2 — dias", 1, 30,
                                  value=CONFIG['ventana_dias_pass2'])
        v_pass3 = st.number_input("Pass 3 — dias", 1, 60,
                                  value=CONFIG['ventana_dias_pass3'])
        v_pass4 = st.number_input("Pass 4 — dias", 1, 180,
                                  value=CONFIG['ventana_dias_pass4'])
        v_dup   = st.number_input("Pass 5 (duplicados) — dias", 1, 60,
                                  value=CONFIG['ventana_dias_duplicados'])
        v_split = st.number_input("Pass 6 (splits) — dias", 1, 60,
                                  value=CONFIG['ventana_dias_split'])
        max_sp  = st.number_input("Max. partes en split", 2, 8,
                                  value=CONFIG['max_split_parts'])

    st.markdown("---")
    st.markdown(
        "<span style='font-size:0.72rem; color:#9ca3af;'>"
        "v1.0 — Motor de conciliación multi-pass</span>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
#  Subida de archivos
# ─────────────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    banco_file = st.file_uploader(
        "Extracto bancario",
        type=["xls", "xlsx"],
        help="Archivo del banco en formato .xls o .xlsx",
    )

with col2:
    mayor_file = st.file_uploader(
        "Libro mayor 572.1",
        type=["xls", "xlsx"],
        help="Libro mayor exportado desde CegidDiezCon",
    )

st.markdown("<div style='height: 0.8rem'></div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  Ejecución
# ─────────────────────────────────────────────────────────────────────────────

if banco_file and mayor_file:

    if st.button("Ejecutar conciliación", use_container_width=True, type="primary"):

        # Aplicar parámetros
        CONFIG['fecha_inicio'] = f'{año}-01-01'
        CONFIG['fecha_fin']    = f'{año}-12-31'
        CONFIG['ventana_dias_pass2'] = v_pass2
        CONFIG['ventana_dias_pass3'] = v_pass3
        CONFIG['ventana_dias_pass4'] = v_pass4
        CONFIG['ventana_dias_duplicados'] = v_dup
        CONFIG['ventana_dias_split'] = v_split
        CONFIG['max_split_parts'] = max_sp

        # Archivos temporales
        tmpdir = Path(tempfile.mkdtemp())
        banco_path = tmpdir / banco_file.name
        mayor_path = tmpdir / mayor_file.name
        banco_path.write_bytes(banco_file.getvalue())
        mayor_path.write_bytes(mayor_file.getvalue())

        try:
            with st.spinner("Cargando extracto bancario..."):
                banco_df = load_banco(banco_path)

            with st.spinner("Cargando libro mayor..."):
                mayor_df, saldo_apertura = load_mayor(mayor_path)

            with st.spinner("Ejecutando conciliación..."):
                engine = MatchingEngine(banco_df, mayor_df)
                engine.run()
                results = engine.get_results()

            with st.spinner("Generando informe..."):
                stem = banco_file.name.upper()
                for word in ['BANCOS', 'BANCO', '572.1', '572_1', '.XLS', '.XLSX']:
                    stem = stem.replace(word, '')
                empresa = stem.strip().strip('_').strip('-').strip() or 'EMPRESA'
                ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_name = f"CONCILIACION_{empresa}_{año}_{ts}.xlsx"

                output_path = tmpdir / output_name
                generate_report(results, saldo_apertura, output_path)
                excel_bytes = output_path.read_bytes()

            # ── Métricas ─────────────────────────────────────────────────
            fecha_inicio = pd.Timestamp(CONFIG['fecha_inicio'])
            fecha_fin    = pd.Timestamp(CONFIG['fecha_fin'])

            banco_p = banco_df[
                (banco_df['Fecha'] >= fecha_inicio) &
                (banco_df['Fecha'] <= fecha_fin)
            ]
            mayor_p = mayor_df[
                (mayor_df['Fecha'] >= fecha_inicio) &
                (mayor_df['Fecha'] <= fecha_fin)
            ]
            matched_b = results['matched_banco'] & set(banco_p['banco_idx'])
            matched_m = results['matched_mayor'] & set(mayor_p['mayor_idx'])

            pct_b = len(matched_b) / max(len(banco_p), 1) * 100
            pct_m = len(matched_m) / max(len(mayor_p), 1) * 100
            pend_b = len(banco_p) - len(matched_b)
            pend_m = len(mayor_p) - len(matched_m)

            n_alta  = sum(1 for m in results['matches'] if m['confianza'] == 'ALTA')
            n_media = sum(1 for m in results['matches'] if m['confianza'] == 'MEDIA')
            n_split = len(results['splits'])

            # ── Resultados ───────────────────────────────────────────────
            st.markdown(
                '<div class="results-ok">Conciliación completada</div>',
                unsafe_allow_html=True,
            )

            c1, c2 = st.columns(2)
            with c1:
                st.metric("Banco conciliado",
                          f"{len(matched_b)} / {len(banco_p)}",
                          f"{pct_b:.1f}%")
            with c2:
                st.metric("Mayor conciliado",
                          f"{len(matched_m)} / {len(mayor_p)}",
                          f"{pct_m:.1f}%")

            c3, c4, c5 = st.columns(3)
            with c3:
                st.metric("Confianza alta", n_alta)
            with c4:
                st.metric("Confianza media", n_media)
            with c5:
                st.metric("Agrupaciones 1:N", n_split)

            st.markdown(
                f'<div class="results-summary">'
                f'Banco pendiente: {pend_b} movimientos &nbsp;&middot;&nbsp; '
                f'Mayor pendiente: {pend_m} apuntes</div>',
                unsafe_allow_html=True,
            )

            # ── Descarga ─────────────────────────────────────────────────
            st.download_button(
                label=f"Descargar {output_name}",
                data=excel_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

else:
    st.info("Sube los dos archivos para comenzar.")
