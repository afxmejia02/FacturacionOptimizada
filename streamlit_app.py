from __future__ import annotations

import importlib.util
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent

# Import the extraction modules in a headless-safe way.
from facturacion import gui_validation_app as validator


def load_payroll_module():
    payroll_path = ROOT / "mapa-de-cargos" / "gui_app.py"
    spec = importlib.util.spec_from_file_location("payroll_module", str(payroll_path))
    if spec is None or spec.loader is None:
        raise RuntimeError("No se pudo cargar el módulo de conciliación de nómina.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


payroll = load_payroll_module()


st.set_page_config(page_title="Pagos y Mapa de Cargos", layout="wide")
st.title("Pagos y Mapa de Cargos")
st.caption("Pagos (Perfiles, Equipos, Servicios) y Mapa de Cargos")


def format_count(value):
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


def style_estado_table(df_display: pd.DataFrame) -> pd.io.formats.style.Styler:
    def row_style(row: pd.Series):
        estado = str(row.get("Estado", "")).strip().lower()
        if estado == "ok":
            color = "#d9f7d9"
            text = "#166534"
        elif "ibc sin soporte" in estado:
            color = "#fff4cc"
            text = "#8a5b00"
        else:
            color = "#ffd6d6"
            text = "#8b1e1e"
        return [f"background-color: {color}; color: {text};"] * len(row)

    return df_display.style.apply(row_style, axis=1)


def build_validation_rows(pdf_path: str, excel_path: str, tipo: str):
    validator_obj = validator.ServicesValidationApp.__new__(validator.ServicesValidationApp)
    rows = []

    if tipo == "perfiles":
        conteo_pdf, fecha = validator_obj._extraer_perfiles_pdf(pdf_path)
        if not conteo_pdf:
            return rows, "No se encontraron perfiles en el PDF."

        conteo_excel = validator_obj._extraer_conteo_excel_perfiles(excel_path, fecha)
        excluded_profiles = {"none", "incapacidad", "observaciones"}
        pdf_perfiles = sorted(
            perfil
            for perfil, cantidad in conteo_pdf.items()
            if str(perfil).strip().lower() not in excluded_profiles and float(cantidad or 0) > 0
        )

        for perfil in pdf_perfiles:
            pdf_cnt = format_count(conteo_pdf.get(perfil, 0))
            excel_cnt = format_count(conteo_excel.get(perfil, 0))
            estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
            rows.append({"Nivel/Perfil": perfil, "PDF": pdf_cnt, "Excel": excel_cnt, "Estado": estado})

        return rows, None

    df_pdf = validator_obj._extraer_conteo_pdf(pdf_path, tipo)
    if df_pdf is None or df_pdf.empty:
        return rows, f"No se encontraron elementos válidos de tipo {tipo} en el PDF."

    pdf_agg = df_pdf.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

    for fecha in sorted(pdf_agg["FECHA"].dropna().unique()):
        conteo_excel = validator_obj._extraer_conteo_excel(excel_path, fecha)
        pdf_fecha = pdf_agg[pdf_agg["FECHA"] == fecha]
        all_servicios = sorted(pdf_fecha["TIPO DE EQUIPO"].dropna().astype(str).unique().tolist())

        for servicio in all_servicios:
            pdf_match = pdf_fecha[pdf_fecha["TIPO DE EQUIPO"] == servicio]
            pdf_cnt = format_count(pdf_match["CANTIDAD"].sum()) if not pdf_match.empty else 0
            if float(pdf_cnt or 0) == 0:
                continue
            excel_cnt = format_count(conteo_excel.get(servicio, 0))
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

    return rows, None


def build_reconciliation_rows(despr_dir: str, trans_dir: str, seg_dir: str, mode: str):
    payroll_obj = payroll.PayrollReconciliationApp.__new__(payroll.PayrollReconciliationApp)
    df_despr = payroll_obj._process_desprendibles(despr_dir)
    df_trans = None
    df_seg = None
    if mode == "transferencias":
        df_trans = payroll_obj._process_transferencia(trans_dir)
    elif mode == "seguridad":
        df_seg = payroll_obj.procesar_seguridad_social(seg_dir) if os.listdir(seg_dir) else None
    df_t, df_s = payroll_obj._reconcile_data(df_despr, df_trans, df_seg)
    if (df_t is None or df_t.empty) and (df_s is None or df_s.empty):
        return pd.DataFrame(), pd.DataFrame(), None

    df_t_display = df_t.copy() if df_t is not None else pd.DataFrame()
    df_s_display = df_s.copy() if df_s is not None else pd.DataFrame()

    def normalize_money_like(v):
        if v is None:
            return ""
        try:
            import numpy as _np
        except Exception:
            _np = None

        if isinstance(v, (list, tuple, set)):
            parts = []
            for x in v:
                try:
                    if _np is not None and isinstance(x, _np.generic):
                        x = int(x)
                    parts.append(str(int(x)))
                except Exception:
                    parts.append(str(x))
            raw_value = " ".join(parts)
        else:
            try:
                if _np is not None and isinstance(v, _np.generic):
                    raw_value = str(int(v))
                elif isinstance(v, (int, float)):
                    raw_value = str(int(v))
                else:
                    raw_value = str(v)
            except Exception:
                raw_value = str(v)

        try:
            return payroll.PayrollReconciliationApp.formatear_valores(None, raw_value)
        except Exception:
            return raw_value

    for column in ["Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"]:
        if column in df_t_display.columns:
            df_t_display[column] = df_t_display[column].apply(normalize_money_like)
        if column in df_s_display.columns:
            df_s_display[column] = df_s_display[column].apply(normalize_money_like)

    return df_t_display, df_s_display, None


mode = st.radio(
    "Tipo de proceso",
    ["Pagos (Perfiles, Equipos, Servicios)", "Mapa de Cargos (Transferencias)", "Mapa de Cargos (Seguridad Social)"],
    horizontal=True,
)

if mode == "Pagos (Perfiles, Equipos, Servicios)":
    tipo = st.selectbox("Tipo de validación", ["equipos", "servicios", "perfiles"])
    pdf_file = st.file_uploader("PDF", type=["pdf"])
    excel_file = st.file_uploader("Excel", type=["xlsx", "xls"])

    if st.button("Procesar validación", type="primary"):
        if not pdf_file or not excel_file:
            st.error("Debes subir un PDF y un archivo Excel.")
        else:
            with tempfile.TemporaryDirectory(prefix="streamlit_validation_") as tmp_dir:
                pdf_path = Path(tmp_dir) / "upload.pdf"
                excel_path = Path(tmp_dir) / "upload.xlsx"
                pdf_path.write_bytes(pdf_file.getbuffer())
                excel_path.write_bytes(excel_file.getbuffer())

                rows, error = build_validation_rows(str(pdf_path), str(excel_path), tipo)
                if error:
                    st.warning(error)
                elif not rows:
                    st.info("No se encontraron datos para comparar.")
                else:
                    df_display = pd.DataFrame(rows)
                    st.dataframe(style_estado_table(df_display), use_container_width=True, hide_index=True)

elif mode == "Mapa de Cargos (Transferencias)":
    despr_files = st.file_uploader("PDFs de desprendibles", type=["pdf"], accept_multiple_files=True)
    trans_files = st.file_uploader("PDFs de transferencias", type=["pdf"], accept_multiple_files=True)

    if st.button("Procesar mapa de cargos", type="primary"):
        if not despr_files or not trans_files:
            st.error("Debes subir al menos un PDF de desprendibles y uno de transferencias.")
        else:
            with tempfile.TemporaryDirectory(prefix="streamlit_mapa_transferencias_") as tmp_dir:
                tmp_base = Path(tmp_dir)
                dir_despr = tmp_base / "despr"
                dir_trans = tmp_base / "trans"
                dir_seg = tmp_base / "seg"
                dir_despr.mkdir(parents=True, exist_ok=True)
                dir_trans.mkdir(parents=True, exist_ok=True)
                dir_seg.mkdir(parents=True, exist_ok=True)

                for file_obj in despr_files:
                    (dir_despr / file_obj.name).write_bytes(file_obj.getbuffer())
                for file_obj in trans_files:
                    (dir_trans / file_obj.name).write_bytes(file_obj.getbuffer())

                try:
                    df_t_display, df_s_display, error = build_reconciliation_rows(
                        str(dir_despr),
                        str(dir_trans),
                        str(dir_seg),
                        "transferencias",
                    )
                    if error:
                        st.warning(error)
                    elif df_t_display is None or df_t_display.empty:
                        st.info("No se encontraron registros o diferencias.")
                    else:
                        st.subheader("Revisión Transferencias")
                        st.dataframe(style_estado_table(df_t_display), use_container_width=True, hide_index=True)
                except Exception as exc:
                    st.error(f"Procesamiento fallido: {exc}")
else:
    despr_files = st.file_uploader("PDFs de desprendibles", type=["pdf"], accept_multiple_files=True)
    seg_files = st.file_uploader("PDFs de seguridad social", type=["pdf"], accept_multiple_files=True)

    if st.button("Procesar mapa de cargos", type="primary"):
        if not despr_files or not seg_files:
            st.error("Debes subir al menos un PDF de desprendibles y uno de seguridad social.")
        else:
            with tempfile.TemporaryDirectory(prefix="streamlit_mapa_seguridad_") as tmp_dir:
                tmp_base = Path(tmp_dir)
                dir_despr = tmp_base / "despr"
                dir_trans = tmp_base / "trans"
                dir_seg = tmp_base / "seg"
                dir_despr.mkdir(parents=True, exist_ok=True)
                dir_trans.mkdir(parents=True, exist_ok=True)
                dir_seg.mkdir(parents=True, exist_ok=True)

                for file_obj in despr_files:
                    (dir_despr / file_obj.name).write_bytes(file_obj.getbuffer())
                for file_obj in seg_files:
                    (dir_seg / file_obj.name).write_bytes(file_obj.getbuffer())

                try:
                    df_t_display, df_s_display, error = build_reconciliation_rows(
                        str(dir_despr),
                        str(dir_trans),
                        str(dir_seg),
                        "seguridad",
                    )
                    if error:
                        st.warning(error)
                    elif df_s_display is None or df_s_display.empty:
                        st.info("No se encontraron registros o diferencias.")
                    else:
                        st.subheader("Revisión Seguridad Social (IBC)")
                        st.dataframe(style_estado_table(df_s_display), use_container_width=True, hide_index=True)
                except Exception as exc:
                    st.error(f"Procesamiento fallido: {exc}")
