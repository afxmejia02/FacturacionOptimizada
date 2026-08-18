"""Streamlit entry point for the validation & reconciliation tool.

This file owns only the presentation/UI flow. The actual extraction and
comparison logic lives in:
  - ``processing.py``    – orchestration (PDF/Excel validation, reconciliation)
  - ``pdf_export.py``    – downloadable PDF with the results
  - ``rendering.py``     – HTML tables and value formatting

Run locally with:  ``streamlit run app.py``
"""
#prueba macos
from __future__ import annotations

import base64
import json
import traceback

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from excel_export import build_results_excel
from pdf_export import build_results_pdf
from processing import process_mano_obra, process_pagos, process_reconciliation
from rendering import build_colored_table, format_number_co

TOOL_LABELS = {
    "pagos": "Pagos Perfiles, Servicios y Equipos",
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


# Plantilla del componente HTML que abre el diálogo nativo "Guardar como" del
# navegador. Los marcadores /*...*/ se sustituyen en _boton_guardar.
_PLANTILLA_GUARDAR = """
<div style="font-family:'Source Sans Pro',sans-serif;">
  <button id="b" style="width:100%;padding:0.45rem 0.75rem;border-radius:0.5rem;
      border:1px solid rgba(49,51,63,0.2);background:#FFFFFF;color:#31333F;
      cursor:pointer;font-size:0.95rem;line-height:1.4;">/*ETIQUETA*/</button>
  <div id="m" style="font-size:0.8rem;margin-top:4px;color:#155724;"></div>
</div>
<script>
(function(){
  const b64="/*B64*/";
  const nombre=/*NOMBRE*/, mime=/*MIME*/, ext=/*EXT*/;
  const bytes=Uint8Array.from(atob(b64), function(c){return c.charCodeAt(0);});
  const b=document.getElementById("b"), m=document.getElementById("m");
  function descargaNormal(){
    const a=document.createElement("a");
    const url=URL.createObjectURL(new Blob([bytes],{type:mime}));
    a.href=url; a.download=nombre; document.body.appendChild(a); a.click();
    a.remove(); URL.revokeObjectURL(url);
    m.style.color="#155724"; m.textContent="Descargado: "+nombre;
  }
  b.addEventListener("click", async function(){
    m.textContent="";
    if(window.showSaveFilePicker){
      try{
        const opts={suggestedName:nombre};
        if(mime && ext){opts.types=[{description:nombre, accept:Object.fromEntries([[mime,[ext]]])}];}
        const h=await window.showSaveFilePicker(opts);
        const w=await h.createWritable();
        await w.write(new Blob([bytes],{type:mime}));
        await w.close();
        m.style.color="#155724"; m.textContent="Guardado \\u2713";
      }catch(e){
        if(e && e.name==="AbortError"){ m.style.color="#888"; m.textContent="Cancelado"; }
        else { descargaNormal(); }
      }
    } else { descargaNormal(); }
  });
})();
</script>
"""


def _boton_guardar(datos: bytes, nombre_sugerido: str, mime: str, etiqueta: str) -> None:
    """Botón que abre el diálogo nativo del navegador "Guardar como".

    Usa la File System Access API (``showSaveFilePicker``) de Edge/Chrome: abre el
    explorador del sistema para que el usuario elija **carpeta y nombre** antes de
    escribir el archivo (no se descarga de inmediato). Si el navegador no soporta
    la API (o la bloquea), cae a una descarga normal con el nombre sugerido.

    Los bytes se incrustan en el componente (base64) porque el clic ocurre dentro
    del iframe (gesto del usuario), no en un rerun de Python.
    """
    b64 = base64.b64encode(datos).decode("ascii")
    ext = "." + nombre_sugerido.rsplit(".", 1)[-1] if "." in nombre_sugerido else ""
    html = (
        _PLANTILLA_GUARDAR
        .replace("/*ETIQUETA*/", etiqueta)
        .replace("/*B64*/", b64)
        .replace("/*NOMBRE*/", json.dumps(nombre_sugerido))
        .replace("/*MIME*/", json.dumps(mime))
        .replace("/*EXT*/", json.dumps(ext))
    )
    components.html(html, height=72)


def _render_export_buttons(key, default_name, titulo, archivos, secciones, source_labels=None) -> None:
    """Botones "Guardar como PDF" / "Guardar como Excel".

    Al hacer clic se abre el diálogo nativo del navegador para elegir carpeta y
    nombre (no se descarga de inmediato con un nombre por defecto).
    """
    st.markdown("**Exportar resultados** — elige carpeta y nombre al guardar")
    col_pdf, col_xlsx = st.columns(2)

    with col_pdf:
        try:
            pdf_bytes = build_results_pdf(titulo, archivos, secciones, source_labels=source_labels)
            _boton_guardar(pdf_bytes, f"{default_name}.pdf", "application/pdf", "Guardar como PDF")
        except Exception as exc:
            print("[ERROR][web_ui] Generación de PDF fallida")
            traceback.print_exc()
            st.error(f"No se pudo generar el PDF: {exc}")

    with col_xlsx:
        try:
            xlsx_bytes = build_results_excel(titulo, archivos, secciones, source_labels=source_labels)
            _boton_guardar(
                xlsx_bytes,
                f"{default_name}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Guardar como Excel",
            )
        except Exception as exc:
            print("[ERROR][web_ui] Generación de Excel fallida")
            traceback.print_exc()
            st.error(f"No se pudo generar el Excel: {exc}")




def _render_pagos_form() -> None:
    tipo = st.selectbox(
        "Tipo de validación",
        ["equipos y servicios", "perfiles"],
    )
    formato = st.selectbox(
        "Formato",
        ["tabarca", "italco"],
        format_func=lambda value: value.upper(),
        help="El formato ITALCO solo aplica a la revisión de perfiles.",
    )
    pdf_files = st.file_uploader(
        "Archivos PDF", type=["pdf"], accept_multiple_files=True
    )
    excel_files = st.file_uploader(
        "Archivos Excel", type=["xlsx", "xls"], accept_multiple_files=True
    )
    submit = st.form_submit_button("Procesar")

    if not submit:
        return

    if not pdf_files:
        st.error("Por favor sube al menos un archivo PDF.")
        return
    if not excel_files:
        st.error("Por favor sube al menos un archivo Excel.")
        return

    try:
        with st.spinner("Procesando archivos..."):
            df_display, message = process_pagos(pdf_files, excel_files, tipo, formato)
        if message:
            st.info(message)
            st.session_state.pagos_result_df = None
            st.session_state.pagos_result_tipo = None
            st.session_state.pagos_files = None
        else:
            st.session_state.pagos_result_df = df_display
            st.session_state.pagos_result_tipo = tipo
            st.session_state.pagos_files = {
                "PDF": ", ".join(f.name for f in pdf_files),
                "Excel": ", ".join(f.name for f in excel_files),
            }
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
    formato = st.selectbox(
        "Formato",
        ["tabarca", "italco"],
        format_func=lambda value: value.upper(),
        help=(
            "ITALCO usa la progresión (filas de DIFERENCIA, nombre completo y "
            "fechas de actividades) en lugar de la hoja 'Informe' de TABARCA."
        ),
    )
    informe_files = st.file_uploader(
        "Excel Informe de Costo",
        type=["xlsx", "xls"],
        key="mano_obra_informe",
        accept_multiple_files=True,
    )
    ods_files = st.file_uploader(
        "Excel ODS (empleados)",
        type=["xlsx", "xls"],
        key="mano_obra_ods",
        accept_multiple_files=True,
    )
    submit = st.form_submit_button("Procesar")
    if not submit:
        return

    if not informe_files:
        st.error("Por favor sube al menos un Excel del Informe de Costo.")
        return
    if not ods_files:
        st.error("Por favor sube al menos un Excel de la ODS.")
        return

    try:
        with st.spinner("Comparando mano de obra..."):
            table_html, message, df_result = process_mano_obra(informe_files, ods_files, formato)
        if message:
            st.info(message)
            st.session_state.mano_obra_df = None
            st.session_state.mano_obra_html = None
            st.session_state.mano_obra_files = None
        else:
            st.session_state.mano_obra_df = df_result
            st.session_state.mano_obra_html = table_html
            st.session_state.mano_obra_files = {
                "Informe de Costo": ", ".join(f.name for f in informe_files),
                "ODS": ", ".join(f.name for f in ods_files),
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

    # Dos filtros independientes: por fecha y por tipo (servicio/equipo/perfil).
    # Se pueden combinar (p. ej. un tipo en todas las fechas) o usar por separado.
    col_fecha, col_tipo = st.columns(2)

    with col_fecha:
        if "Fecha" in df_pagos.columns:
            available_dates = sorted(df_pagos["Fecha"].dropna().astype(str).unique().tolist())
            selected_date = st.selectbox(
                "Filtrar por fecha",
                ["Todas las fechas"] + available_dates,
                key="pagos_date_filter",
            )
            if selected_date != "Todas las fechas":
                df_pagos = df_pagos[df_pagos["Fecha"].astype(str) == selected_date].copy()

    # La columna de "tipo" se llama "Servicio" (equipos/servicios) o
    # "Nivel/Perfil" (perfiles), según la validación.
    columna_tipo = next(
        (c for c in ("Servicio", "Nivel/Perfil") if c in df_pagos.columns), None
    )
    with col_tipo:
        if columna_tipo:
            tipos_disponibles = sorted(df_pagos[columna_tipo].dropna().astype(str).unique().tolist())
            etiqueta = "tipo" if columna_tipo == "Servicio" else columna_tipo.lower()
            selected_tipo = st.selectbox(
                f"Filtrar por {etiqueta}",
                ["Todos"] + tipos_disponibles,
                key="pagos_tipo_filter",
            )
            if selected_tipo != "Todos":
                df_pagos = df_pagos[df_pagos[columna_tipo].astype(str) == selected_tipo].copy()

    # Columnas PDF/Excel en formato colombiano (punto de miles, coma decimal).
    def _formato_co(df):
        out = df.copy()
        for col in ("PDF", "Excel"):
            if col in out.columns:
                out[col] = out[col].map(format_number_co)
        return out

    st.markdown(build_colored_table(_formato_co(df_pagos)), unsafe_allow_html=True)

    tipo = st.session_state.get("pagos_result_tipo") or "pagos"
    _render_export_buttons(
        key="pagos",
        default_name=f"resultado_{tipo}",
        titulo=f"Validación de {tipo}",
        archivos=st.session_state.get("pagos_files"),
        secciones=[("", _formato_co(df_full))],  # el PDF incluye todas las fechas, no solo la filtrada
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
