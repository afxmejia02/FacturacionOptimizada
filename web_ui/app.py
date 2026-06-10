"""Streamlit entry point for the validation & reconciliation tool.

This file owns only the presentation/UI flow. The actual extraction and
comparison logic lives in:
  - ``processing.py``    – orchestration (PDF/Excel validation, reconciliation)
  - ``pdf_export.py``    – downloadable PDF with the results
  - ``rendering.py``     – HTML tables and value formatting

Run locally with:  ``streamlit run app.py``
"""
from __future__ import annotations

import os
import traceback

import pandas as pd
import streamlit as st

from excel_export import build_results_excel
from pdf_export import build_results_pdf
from processing import process_mano_obra, process_pagos, process_reconciliation
from rendering import build_colored_table

TOOL_LABELS = {
    "pagos": "Validación PDF + Excel",
    "mapa_transferencias": "Mapa de cargos - transferencias",
    "mapa_seguridad": "Mapa de cargos - seguridad social",
    "mapa_mano_obra": "Mapa de cargos - mano de obra",
}

_RESULT_KEYS = {
    "pagos_result_df": None,
    "pagos_result_tipo": None,
    "pagos_files": None,
    "recon_transfer_df": None,
    "recon_seguridad_df": None,
    "recon_mode": None,
    "recon_parts": [],
    "recon_tables": [],
    "recon_files": None,
    "mano_obra_df": None,
    "mano_obra_html": None,
    "mano_obra_files": None,
}


def _init_session_state() -> None:
    defaults = {**_RESULT_KEYS, "active_tool": None}
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _reset_results() -> None:
    """Clear every stored result so one tool's output never lingers under another."""
    for key, value in _RESULT_KEYS.items():
        st.session_state[key] = value


def _dialogo_guardar(nombre_sugerido, defaultextension, filetypes):
    """Abre el diálogo nativo "Guardar como" del sistema operativo.

    El servidor de Streamlit corre en la máquina del usuario (localhost), así que
    el diálogo de tkinter aparece en su pantalla y permite elegir ruta y nombre.
    Se ejecuta en un **subproceso** (hilo principal limpio) porque crear tkinter
    dentro del hilo de trabajo de Streamlit es inestable en Windows.

    Devuelve la ruta elegida, ``""`` si el usuario cancela, o ``None`` si no hay
    GUI nativa disponible (entorno headless) -> el llamador usa descarga por
    navegador como respaldo.
    """
    import json
    import subprocess
    import sys

    # Script hijo: abre el diálogo y escribe la ruta (UTF-8) en stdout.
    script = (
        "import sys, json\n"
        "import tkinter as tk\n"
        "from tkinter import filedialog\n"
        "args = json.loads(sys.argv[1])\n"
        "root = tk.Tk(); root.withdraw()\n"
        "try:\n"
        "    root.attributes('-topmost', True)\n"
        "except Exception:\n"
        "    pass\n"
        "ruta = filedialog.asksaveasfilename(initialfile=args['nombre'],"
        " defaultextension=args['ext'], filetypes=[tuple(ft) for ft in args['filetypes']])\n"
        "sys.stdout.buffer.write((ruta or '').encode('utf-8'))\n"
    )
    payload = json.dumps(
        {"nombre": nombre_sugerido, "ext": defaultextension, "filetypes": filetypes}
    )
    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        res = subprocess.run(
            [sys.executable, "-c", script, payload],
            capture_output=True,
            timeout=300,
            **kwargs,
        )
    except Exception:
        print("[ERROR][web_ui] El diálogo nativo de guardado falló")
        traceback.print_exc()
        return None

    if res.returncode != 0:
        # tkinter no disponible u otro error en el hijo -> respaldo por navegador.
        if res.stderr:
            print("[web_ui] diálogo (subproceso) stderr:", res.stderr.decode("utf-8", "ignore").strip())
        return None
    return res.stdout.decode("utf-8", "ignore").strip("\r\n")  # "" cuando se cancela


def _guardar_con_dialogo(construir_bytes, nombre_sugerido, defaultextension, filetypes, key, etiqueta):
    """Genera los bytes y los guarda en la ruta que el usuario elija.

    Abre el diálogo nativo del sistema; si no está disponible (p. ej. en un
    entorno sin GUI) deja preparada una descarga por navegador como respaldo.
    """
    try:
        datos = construir_bytes()
    except Exception as exc:
        print(f"[ERROR][web_ui] Generación de {etiqueta} fallida")
        traceback.print_exc()
        st.error(f"No se pudo generar el {etiqueta}: {exc}")
        return

    ruta = _dialogo_guardar(nombre_sugerido, defaultextension, filetypes)
    if ruta is None:
        # Sin diálogo nativo: respaldo por descarga del navegador (persistente).
        st.session_state[f"_pending_dl_{key}_{etiqueta}"] = (datos, nombre_sugerido)
        return
    if ruta == "":
        st.info("Guardado cancelado.")
        return
    try:
        with open(ruta, "wb") as fh:
            fh.write(datos)
        st.success(f"Guardado en: {ruta}")
    except Exception as exc:
        st.error(f"No se pudo guardar el archivo: {exc}")


def _render_export_buttons(key, default_name, titulo, archivos, secciones, source_labels=None) -> None:
    """Botones "Guardar como PDF" / "Guardar como Excel".

    Al hacer clic se generan los bytes y se abre el diálogo nativo del sistema
    para elegir ruta y nombre (no se descarga de inmediato).
    """
    st.markdown("**Exportar resultados**")
    col_pdf, col_xlsx = st.columns(2)

    with col_pdf:
        if st.button("Guardar como PDF", key=f"btn_pdf_{key}"):
            _guardar_con_dialogo(
                lambda: build_results_pdf(titulo, archivos, secciones, source_labels=source_labels),
                f"{default_name}.pdf",
                ".pdf",
                [("Archivo PDF", "*.pdf"), ("Todos los archivos", "*.*")],
                key,
                "PDF",
            )
    with col_xlsx:
        if st.button("Guardar como Excel", key=f"btn_xlsx_{key}"):
            _guardar_con_dialogo(
                lambda: build_results_excel(titulo, archivos, secciones, source_labels=source_labels),
                f"{default_name}.xlsx",
                ".xlsx",
                [("Libro de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
                key,
                "Excel",
            )

    # Respaldo: si el diálogo nativo no estaba disponible, ofrecer descarga.
    for etiqueta, mime in (("PDF", "application/pdf"),
                           ("Excel", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")):
        pendiente = st.session_state.get(f"_pending_dl_{key}_{etiqueta}")
        if pendiente:
            datos, nombre = pendiente
            st.download_button(
                label=f"Descargar {etiqueta}",
                data=datos,
                file_name=nombre,
                mime=mime,
                key=f"dlbtn_{key}_{etiqueta}",
            )




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
            df_display, message = process_pagos(pdf_file, excel_file, tipo)
        if message:
            st.info(message)
            st.session_state.pagos_result_df = None
            st.session_state.pagos_result_tipo = None
            st.session_state.pagos_files = None
        else:
            st.session_state.pagos_result_df = df_display
            st.session_state.pagos_result_tipo = tipo
            st.session_state.pagos_files = {"PDF": pdf_file.name, "Excel": excel_file.name}
    except Exception as exc:
        print("[ERROR][web_ui] Procesamiento fallido en pagos")
        traceback.print_exc()
        st.session_state.pagos_result_df = None
        st.session_state.pagos_result_tipo = None
        st.error(f"Procesamiento fallido: {exc}")


def _render_reconciliation_form(app_choice: str) -> None:
    recon_mode = "transfers" if app_choice == "mapa_transferencias" else "seguridad"

    format_label = (
        "Formato de transferencias" if recon_mode == "transfers" else "Formato de seguridad social"
    )
    transfer_format = st.selectbox(
        format_label,
        ["tabarca", "italco"],
        format_func=lambda value: value.upper(),
    )

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
            parts, message, df_transfers, df_seguridad, tables = process_reconciliation(
                despr_files, trans_files, seguridad_files, recon_mode, transfer_format
            )
            st.session_state.recon_transfer_df = df_transfers
            st.session_state.recon_seguridad_df = df_seguridad
            st.session_state.recon_mode = recon_mode
        if message:
            st.info(message)
            st.session_state.recon_parts = []
            st.session_state.recon_tables = []
            st.session_state.recon_files = None
        else:
            st.session_state.recon_parts = parts
            st.session_state.recon_tables = tables
            entrada_extra = trans_files if recon_mode == "transfers" else seguridad_files
            etiqueta_extra = "Transferencias" if recon_mode == "transfers" else "Seguridad social"
            st.session_state.recon_files = {
                "Desprendibles": [f.name for f in despr_files],
                etiqueta_extra: [f.name for f in entrada_extra],
            }
    except Exception as exc:
        print("[ERROR][web_ui] Procesamiento de conciliacion fallido")
        traceback.print_exc()
        st.error(f"Procesamiento de conciliación falló: {exc}")
        st.session_state.recon_parts = []
        st.session_state.recon_tables = []
        st.session_state.recon_files = None
        st.session_state.recon_transfer_df = None
        st.session_state.recon_seguridad_df = None
        st.session_state.recon_mode = None


def _render_mano_obra_form() -> None:
    st.caption(
        "Cruza el **Informe de Costo** contra el registro de la **ODS** por número "
        "de documento. Resalta solo la celda del campo que no coincide."
    )
    informe_file = st.file_uploader(
        "Excel Informe de Costo", type=["xlsx", "xls"], key="mano_obra_informe"
    )
    ods_file = st.file_uploader(
        "Excel ODS (empleados)", type=["xlsx", "xls"], key="mano_obra_ods"
    )
    submit = st.form_submit_button("Procesar")
    if not submit:
        return

    if not informe_file:
        st.error("Por favor sube el Excel del Informe de Costo.")
        return
    if not ods_file:
        st.error("Por favor sube el Excel de la ODS.")
        return

    try:
        with st.spinner("Comparando mano de obra..."):
            table_html, message, df_result = process_mano_obra(informe_file, ods_file)
        if message:
            st.info(message)
            st.session_state.mano_obra_df = None
            st.session_state.mano_obra_html = None
            st.session_state.mano_obra_files = None
        else:
            st.session_state.mano_obra_df = df_result
            st.session_state.mano_obra_html = table_html
            st.session_state.mano_obra_files = {
                "Informe de Costo": informe_file.name,
                "ODS": ods_file.name,
            }
    except Exception as exc:
        print("[ERROR][web_ui] Procesamiento de mano de obra fallido")
        traceback.print_exc()
        st.error(f"Procesamiento de mano de obra falló: {exc}")
        st.session_state.mano_obra_df = None
        st.session_state.mano_obra_html = None


def _render_pagos_results() -> None:
    df_full = st.session_state.get("pagos_result_df")
    if not isinstance(df_full, pd.DataFrame) or df_full.empty:
        return

    st.subheader("Resultados por fecha")
    df_pagos = df_full
    if "Fecha" in df_pagos.columns:
        available_dates = sorted(df_pagos["Fecha"].dropna().astype(str).unique().tolist())
        selected_date = st.selectbox(
            "Filtrar por fecha",
            ["Todas las fechas"] + available_dates,
            key="pagos_date_filter",
        )
        if selected_date != "Todas las fechas":
            df_pagos = df_pagos[df_pagos["Fecha"].astype(str) == selected_date].copy()

    st.markdown(build_colored_table(df_pagos), unsafe_allow_html=True)

    tipo = st.session_state.get("pagos_result_tipo") or "pagos"
    _render_export_buttons(
        key="pagos",
        default_name=f"resultado_{tipo}",
        titulo=f"Validación de {tipo}",
        archivos=st.session_state.get("pagos_files"),
        secciones=[("", df_full)],  # el PDF incluye todas las fechas, no solo la filtrada
    )


def _render_reconciliation_results() -> None:
    if not st.session_state.recon_mode or not st.session_state.recon_parts:
        return

    for title, table_html in st.session_state.recon_parts:
        st.subheader(title)
        st.markdown(table_html, unsafe_allow_html=True)

    nombre_base = (
        "reconciliacion_transferencias"
        if st.session_state.recon_mode == "transfers"
        else "reconciliacion_seguridad_social"
    )
    _render_export_buttons(
        key=f"recon_{st.session_state.recon_mode}",
        default_name=nombre_base,
        titulo="Conciliación de nómina",
        archivos=st.session_state.get("recon_files"),
        secciones=st.session_state.get("recon_tables") or [],
    )


def _render_mano_obra_results() -> None:
    table_html = st.session_state.get("mano_obra_html")
    df_result = st.session_state.get("mano_obra_df")
    if not table_html or df_result is None:
        return

    st.subheader("Validación mano de obra")
    st.markdown(table_html, unsafe_allow_html=True)

    _render_export_buttons(
        key="mano_obra",
        default_name="validacion_mano_obra",
        titulo="Validación mano de obra",
        archivos=st.session_state.get("mano_obra_files"),
        secciones=[("", df_result)],
        source_labels=("Informe", "Lista ODS"),
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

    # Al cambiar de herramienta, descartar resultados previos (la tabla de una
    # revisión no debe quedar visible bajo otra).
    if app_choice != st.session_state.active_tool:
        _reset_results()
        st.session_state.active_tool = app_choice

    with st.form("validation_form", clear_on_submit=False):
        if app_choice == "pagos":
            _render_pagos_form()
        elif app_choice == "mapa_mano_obra":
            _render_mano_obra_form()
        else:
            _render_reconciliation_form(app_choice)

    if app_choice == "pagos":
        _render_pagos_results()
    elif app_choice == "mapa_mano_obra":
        _render_mano_obra_results()
    else:
        _render_reconciliation_results()


if __name__ == "__main__":
    main()
