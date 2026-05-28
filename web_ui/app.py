"""Streamlit web UI that reuses functions from `facturacion/gui_validation_app.py`.

The page keeps the same processing logic as the original app entrypoint and only
changes the presentation layer so the project can be shared without running a local
desktop GUI or Flask server.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

# Ensure parent workspace is on sys.path so we can import the existing module
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from facturacion import gui_validation_app as validator


def _build_colored_table(df_display: pd.DataFrame) -> str:
    headers = list(df_display.columns)
    parts = []
    parts.append('<table class="table table-striped table-bordered">')
    parts.append('<thead><tr>')
    for header in headers:
        parts.append(f'<th>{header}</th>')
    parts.append('</tr></thead>')
    parts.append('<tbody>')

    for _, row in df_display.iterrows():
        estado = str(row.get("Estado", "")).strip().lower()
        if estado == "ok":
            row_class = "table-success"
        elif "ibc sin soporte" in estado:
            row_class = "table-warning"
        else:
            row_class = "table-danger"
        parts.append(f'<tr class="{row_class}">')
        for header in headers:
            value = row[header] if pd.notna(row[header]) else ""
            parts.append(f'<td>{value}</td>')
        parts.append('</tr>')

    parts.append('</tbody></table>')
    return "".join(parts)


def _format_count(value):
    if value is None:
        return 0
    try:
        import numpy as _np
    except Exception:
        _np = None

    if _np is not None and isinstance(value, _np.generic):
        value = value.item()
    try:
        numeric = float(value)
        if numeric.is_integer():
            return int(numeric)
        return numeric
    except Exception:
        return value


def _normalize_list_like(value):
    if value is None:
        return ""
    try:
        import numpy as _np
    except Exception:
        _np = None

    if isinstance(value, (list, tuple, set)):
        parts = []
        for item in value:
            try:
                if _np is not None and isinstance(item, _np.generic):
                    item = int(item)
                parts.append(str(int(item)))
            except Exception:
                parts.append(str(item))
        return " ".join(parts)

    try:
        if _np is not None and isinstance(value, _np.generic):
            return str(int(value))
        if isinstance(value, (int, float)):
            return str(int(value))
    except Exception:
        pass

    return str(value)


def _format_dataframe(df_in: pd.DataFrame, formatter) -> pd.DataFrame:
    if df_in is None or df_in.empty:
        return df_in
    df_tmp = df_in.copy()
    list_cols = [col for col in ["Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"] if col in df_tmp.columns]

    for column in list_cols:
        def _fmt_cell(cell):
            normalized = _normalize_list_like(cell)
            try:
                return formatter(normalized)
            except Exception:
                return normalized

        df_tmp[column] = df_tmp[column].apply(_fmt_cell)
    return df_tmp


def _load_payroll_module():
    payroll_path = ROOT / "mapa-de-cargos" / "gui_app.py"
    if not payroll_path.exists():
        raise FileNotFoundError("No se encontró el módulo de conciliación (mapa-de-cargos/gui_app.py).")

    spec = importlib.util.spec_from_file_location("payroll_module", str(payroll_path))
    payroll = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(payroll)
    return payroll


def _process_pagos(pdf_file, excel_file, tipo):
    validator_obj = validator.ServicesValidationApp.__new__(validator.ServicesValidationApp)

    with tempfile.TemporaryDirectory(prefix="web_ui_") as tmp_dir:
        pdf_path = os.path.join(tmp_dir, "upload.pdf")
        excel_path = os.path.join(tmp_dir, "upload.xlsx")

        with open(pdf_path, "wb") as pdf_handle:
            pdf_handle.write(pdf_file.getbuffer())
        with open(excel_path, "wb") as excel_handle:
            excel_handle.write(excel_file.getbuffer())

        rows = []

        if tipo == "perfiles":
            conteo_pdf, fecha = validator_obj._extraer_perfiles_pdf(pdf_path)
            if not conteo_pdf:
                return None, "No se encontraron perfiles en el PDF."

            conteo_excel = validator_obj._extraer_conteo_excel_perfiles(excel_path, fecha)
            excluded_profiles = {"none", "incapacidad", "observaciones"}
            pdf_perfiles = sorted(
                perfil
                for perfil, cantidad in conteo_pdf.items()
                if str(perfil).strip().lower() not in excluded_profiles and float(cantidad or 0) > 0
            )

            for perfil in pdf_perfiles:
                pdf_raw = conteo_pdf.get(perfil, 0)
                try:
                    if float(pdf_raw or 0) == 0:
                        continue
                except Exception:
                    pass

                pdf_cnt = _format_count(pdf_raw)
                excel_cnt = _format_count(conteo_excel.get(perfil, 0))
                estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
                rows.append(
                    {
                        "Nivel/Perfil": perfil,
                        "PDF": pdf_cnt,
                        "Excel": excel_cnt,
                        "Estado": estado,
                    }
                )
        else:
            df_pdf = validator_obj._extraer_conteo_pdf(pdf_path, tipo)
            if df_pdf is None or df_pdf.empty:
                return None, f"No se encontraron elementos válidos de tipo {tipo} en el PDF."

            pdf_agg = df_pdf.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

            for fecha in sorted(pdf_agg["FECHA"].dropna().unique()):
                conteo_excel = validator_obj._extraer_conteo_excel(excel_path, fecha)
                pdf_fecha = pdf_agg[pdf_agg["FECHA"] == fecha]
                all_servicios = sorted(pdf_fecha["TIPO DE EQUIPO"].dropna().astype(str).unique().tolist())

                for servicio in all_servicios:
                    pdf_match = pdf_fecha[pdf_fecha["TIPO DE EQUIPO"] == servicio]
                    pdf_cnt = _format_count(pdf_match["CANTIDAD"].sum()) if not pdf_match.empty else 0
                    if float(pdf_cnt or 0) == 0:
                        continue
                    excel_cnt = _format_count(conteo_excel.get(servicio, 0))
                    estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
                    rows.append(
                        {
                            "Fecha": fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else str(fecha),
                            "Servicio": servicio,
                            "PDF": pdf_cnt,
                            "Excel": excel_cnt,
                            "Estado": estado,
                        }
                    )

        df_display = pd.DataFrame(rows)
        if df_display.empty:
            return None, "No se encontraron datos para comparar."

        table_html = _build_colored_table(df_display)
        return table_html, None


def _process_reconciliation(despr_files, trans_files, seguridad_files, recon_mode):
    payroll = _load_payroll_module()
    PayrollApp = payroll.PayrollReconciliationApp
    payroll_obj = PayrollApp.__new__(PayrollApp)

    with tempfile.TemporaryDirectory(prefix="web_ui_rec_") as tmp_base:
        dir_despr = os.path.join(tmp_base, "despr")
        dir_trans = os.path.join(tmp_base, "trans")
        dir_seg = os.path.join(tmp_base, "seg")
        os.makedirs(dir_despr, exist_ok=True)
        os.makedirs(dir_trans, exist_ok=True)
        os.makedirs(dir_seg, exist_ok=True)

        for uploaded_file in despr_files:
            if uploaded_file and uploaded_file.name:
                with open(os.path.join(dir_despr, uploaded_file.name), "wb") as handle:
                    handle.write(uploaded_file.getbuffer())

        if recon_mode == "transfers":
            for uploaded_file in trans_files:
                if uploaded_file and uploaded_file.name:
                    with open(os.path.join(dir_trans, uploaded_file.name), "wb") as handle:
                        handle.write(uploaded_file.getbuffer())
        else:
            for uploaded_file in seguridad_files:
                if uploaded_file and uploaded_file.name:
                    with open(os.path.join(dir_seg, uploaded_file.name), "wb") as handle:
                        handle.write(uploaded_file.getbuffer())

        df_despr = payroll_obj._process_desprendibles(dir_despr)
        df_trans = payroll_obj._process_transferencia(dir_trans) if os.listdir(dir_trans) else None
        df_seg = payroll_obj.procesar_seguridad_social(dir_seg) if os.listdir(dir_seg) else None

        df_transfers, df_seguridad = payroll_obj._reconcile_data(df_despr, df_trans, df_seg)
        if (df_transfers is None or df_transfers.empty) and (df_seguridad is None or df_seguridad.empty):
            return [], "No se encontraron registros o diferencias."

        df_display_t = df_transfers.copy() if df_transfers is not None else pd.DataFrame()
        df_display_s = df_seguridad.copy() if df_seguridad is not None else pd.DataFrame()

        df_display_t = _format_dataframe(df_display_t, payroll_obj.formatear_valores)
        df_display_s = _format_dataframe(df_display_s, payroll_obj.formatear_valores)

        parts = []
        if recon_mode == "transfers" and df_display_t is not None and not df_display_t.empty:
            parts.append(("Revisión Transferencias", _build_colored_table(df_display_t)))

        if recon_mode == "seguridad" and df_display_s is not None and not df_display_s.empty:
            parts.append(("Revisión Seguridad Social (IBC)", _build_colored_table(df_display_s)))

        if not parts:
            return [], "No se encontraron registros para el modo seleccionado."

        return parts, None


def main():
    st.set_page_config(page_title="Web UI", layout="wide")

    st.title("Sistema de validación y conciliación")
    st.write("Carga tus archivos para comparar PDF contra Excel o para conciliar desprendibles, transferencias y seguridad social.")

    app_choice = st.radio(
        "Selecciona la herramienta",
        ["pagos", "mapa_transferencias", "mapa_seguridad"],
        format_func=lambda value: {
            "pagos": "Validación PDF + Excel",
            "mapa_transferencias": "Mapa de cargos - transferencias",
            "mapa_seguridad": "Mapa de cargos - seguridad social",
        }[value],
        horizontal=True,
    )

    with st.form("validation_form", clear_on_submit=False):
        if app_choice == "pagos":
            tipo = st.selectbox("Tipo de validación", ["equipos", "servicios", "perfiles"])
            pdf_file = st.file_uploader("Archivo PDF", type=["pdf"])
            excel_file = st.file_uploader("Archivo Excel", type=["xlsx", "xls"])
            submit = st.form_submit_button("Procesar")

            if submit:
                if not pdf_file:
                    st.error("Por favor sube un archivo PDF.")
                    return
                if not excel_file:
                    st.error("Por favor sube un archivo Excel.")
                    return

                try:
                    with st.spinner("Procesando archivos..."):
                        table_html, message = _process_pagos(pdf_file, excel_file, tipo)
                    if message:
                        st.info(message)
                    elif table_html:
                        st.markdown(table_html, unsafe_allow_html=True)
                except Exception as exc:
                    st.error(f"Procesamiento fallido: {exc}")

        else:
            recon_mode = "transfers" if app_choice == "mapa_transferencias" else "seguridad"
            despr_files = st.file_uploader(
                "PDFs de desprendibles",
                type=["pdf"],
                accept_multiple_files=True,
            )
            if recon_mode == "transfers":
                trans_files = st.file_uploader(
                    "PDFs de transferencias",
                    type=["pdf"],
                    accept_multiple_files=True,
                )
                seguridad_files = []
            else:
                seguridad_files = st.file_uploader(
                    "PDFs de seguridad social",
                    type=["pdf"],
                    accept_multiple_files=True,
                )
                trans_files = []

            submit = st.form_submit_button("Procesar")

            if submit:
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
                        parts, message = _process_reconciliation(despr_files, trans_files, seguridad_files, recon_mode)
                    if message:
                        st.info(message)
                    else:
                        for title, table_html in parts:
                            st.subheader(title)
                            st.markdown(table_html, unsafe_allow_html=True)
                except Exception as exc:
                    st.error(f"Procesamiento de conciliación falló: {exc}")


if __name__ == "__main__":
    main()
