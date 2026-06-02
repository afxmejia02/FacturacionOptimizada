"""Streamlit entry point for the validation & reconciliation tool.

This file owns only the presentation/UI flow. The actual extraction and
comparison logic lives in:
  - ``processing.py``    – orchestration (PDF/Excel validation, reconciliation)
  - ``excel_export.py``  – downloadable Excel workbook
  - ``rendering.py``     – HTML tables and value formatting

Run locally with:  ``streamlit run app.py``
"""
from __future__ import annotations

import traceback

import pandas as pd
import streamlit as st

from excel_export import build_reconciliation_excel_bytes
from processing import build_perfiles_table, process_pagos, process_reconciliation
from rendering import build_colored_table

TOOL_LABELS = {
    "pagos": "Validación PDF + Excel",
    "mapa_transferencias": "Mapa de cargos - transferencias",
    "mapa_seguridad": "Mapa de cargos - seguridad social",
}


def _init_session_state() -> None:
    defaults = {
        "perfiles_result_df": None,
        "perfiles_result_ready": False,
        "recon_transfer_df": None,
        "recon_seguridad_df": None,
        "recon_mode": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _render_pagos_form() -> None:
    tipo = st.selectbox("Tipo de validación", ["equipos", "servicios", "perfiles"])
    pdf_file = st.file_uploader("Archivo PDF", type=["pdf"])
    excel_file = st.file_uploader("Archivo Excel", type=["xlsx", "xls"])
    submit = st.form_submit_button("Procesar")

    if not submit:
        return

    if not pdf_file:
        st.error("Por favor sube un archivo PDF.")
        return
    if not excel_file:
        st.error("Por favor sube un archivo Excel.")
        return

    try:
        with st.spinner("Procesando archivos..."):
            table_html, message = process_pagos(pdf_file, excel_file, tipo)
        if message:
            st.info(message)
            st.session_state.perfiles_result_df = None
            st.session_state.perfiles_result_ready = False
        elif table_html and tipo != "perfiles":
            st.markdown(table_html, unsafe_allow_html=True)
        elif tipo != "perfiles":
            st.session_state.perfiles_result_df = None
            st.session_state.perfiles_result_ready = False
    except Exception as exc:
        print("[ERROR][web_ui] Procesamiento fallido en pagos")
        traceback.print_exc()
        st.session_state.perfiles_result_df = None
        st.session_state.perfiles_result_ready = False
        st.error(f"Procesamiento fallido: {exc}")

    if tipo == "perfiles" and pdf_file and excel_file:
        try:
            with st.spinner("Construyendo tabla por fechas..."):
                st.session_state.perfiles_result_df = build_perfiles_table(pdf_file, excel_file)
                st.session_state.perfiles_result_ready = True
        except Exception as exc:
            print("[ERROR][web_ui] Construccion de tabla por fecha fallida")
            traceback.print_exc()
            st.session_state.perfiles_result_df = None
            st.session_state.perfiles_result_ready = False
            st.error(f"No se pudo construir la tabla por fecha: {exc}")


def _render_reconciliation_form(app_choice: str) -> None:
    recon_mode = "transfers" if app_choice == "mapa_transferencias" else "seguridad"

    if recon_mode == "transfers":
        transfer_format = st.selectbox(
            "Formato de transferencias",
            ["tabarca", "italco"],
            format_func=lambda value: value.upper(),
        )
    else:
        transfer_format = "tabarca"

    despr_files = st.file_uploader("PDFs de desprendibles", type=["pdf"], accept_multiple_files=True)
    if recon_mode == "transfers":
        trans_files = st.file_uploader("PDFs de transferencias", type=["pdf"], accept_multiple_files=True)
        seguridad_files = []
    else:
        seguridad_files = st.file_uploader("PDFs de seguridad social", type=["pdf"], accept_multiple_files=True)
        trans_files = []

    submit = st.form_submit_button("Procesar")
    if not submit:
        return

    if not despr_files:
        st.error("Por favor sube al menos un PDF de desprendibles.")
        return
    if recon_mode == "transfers" and not trans_files:
        st.error("Por favor sube al menos un PDF de transferencias para el modo 'transferencias'.")
        return
    if recon_mode == "seguridad" and not seguridad_files:
        st.error("Por favor sube al menos un PDF de seguridad social para el modo 'seguridad'.")
        return

    try:
        with st.spinner("Procesando conciliación..."):
            parts, message, df_transfers, df_seguridad = process_reconciliation(
                despr_files, trans_files, seguridad_files, recon_mode, transfer_format
            )
            st.session_state.recon_transfer_df = df_transfers
            st.session_state.recon_seguridad_df = df_seguridad
            st.session_state.recon_mode = recon_mode
        if message:
            st.info(message)
            st.session_state.recon_parts = []
        else:
            st.session_state.recon_parts = parts
    except Exception as exc:
        print("[ERROR][web_ui] Procesamiento de conciliacion fallido")
        traceback.print_exc()
        st.error(f"Procesamiento de conciliación falló: {exc}")
        st.session_state.recon_parts = []
        st.session_state.recon_transfer_df = None
        st.session_state.recon_seguridad_df = None
        st.session_state.recon_mode = None


def _render_perfiles_results() -> None:
    if not (
        st.session_state.perfiles_result_ready
        and isinstance(st.session_state.perfiles_result_df, pd.DataFrame)
    ):
        return

    df_perfiles = st.session_state.perfiles_result_df
    if df_perfiles.empty:
        return

    st.subheader("Resultados por fecha")
    available_dates = sorted(df_perfiles["Fecha"].dropna().astype(str).unique().tolist())
    selected_date = st.selectbox(
        "Filtrar por fecha",
        ["Todas las fechas"] + available_dates,
        key="perfiles_date_filter",
    )

    if selected_date != "Todas las fechas":
        df_perfiles = df_perfiles[df_perfiles["Fecha"].astype(str) == selected_date].copy()

    st.markdown(build_colored_table(df_perfiles), unsafe_allow_html=True)


def _render_reconciliation_results() -> None:
    if "recon_parts" not in st.session_state or not st.session_state.recon_mode:
        return
    if not st.session_state.recon_parts:
        return

    for title, table_html in st.session_state.recon_parts:
        st.subheader(title)
        st.markdown(table_html, unsafe_allow_html=True)

    excel_bytes = build_reconciliation_excel_bytes(
        st.session_state.recon_transfer_df,
        st.session_state.recon_seguridad_df,
        st.session_state.recon_mode,
    )
    file_name = (
        "reconciliacion_transferencias.xlsx"
        if st.session_state.recon_mode == "transfers"
        else "reconciliacion_seguridad_social.xlsx"
    )
    st.download_button(
        label="Descargar Excel con resultados",
        data=excel_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_recon_{st.session_state.recon_mode}",
    )


def main() -> None:
    st.set_page_config(page_title="Validación y conciliación", layout="wide")
    st.title("Sistema de validación y conciliación")
    st.write(
        "Carga tus archivos para comparar PDF contra Excel o para conciliar "
        "desprendibles, transferencias y seguridad social."
    )

    _init_session_state()

    app_choice = st.radio(
        "Selecciona la herramienta",
        list(TOOL_LABELS.keys()),
        format_func=lambda value: TOOL_LABELS[value],
        horizontal=True,
    )

    with st.form("validation_form", clear_on_submit=False):
        if app_choice == "pagos":
            _render_pagos_form()
        else:
            _render_reconciliation_form(app_choice)

    _render_perfiles_results()
    _render_reconciliation_results()


if __name__ == "__main__":
    main()
