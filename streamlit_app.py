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
st.caption("Validación de PDFs contra Excel y conciliación de nómina")


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


def build_colored_table(df_display: pd.DataFrame) -> str:
    headers = list(df_display.columns)
    parts = []
    parts.append('<table class="table table-striped table-bordered">')
    parts.append('<thead><tr>')
    for h in headers:
        parts.append(f"<th>{h}</th>")
    parts.append("</tr></thead>")
    parts.append("<tbody>")

    for _, row in df_display.iterrows():
        estado = str(row.get("Estado", "")).strip().lower()
        if estado == "ok":
            row_class = "table-success"
        elif "ibc sin soporte" in estado:
            row_class = "table-warning"
        else:
            row_class = "table-danger"
        parts.append(f'<tr class="{row_class}">')
        for h in headers:
            val = row[h] if pd.notna(row[h]) else ""
            parts.append(f"<td>{val}</td>")
        parts.append("</tr>")

    parts.append("</tbody></table>")
    return "".join(parts)


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
            pdf_raw = conteo_pdf.get(perfil, 0)
            try:
                if float(pdf_raw or 0) == 0:
                    continue
            except Exception:
                pass

            pdf_cnt = format_count(pdf_raw)
            excel_cnt = format_count(conteo_excel.get(perfil, 0))
            estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
            rows.append(
                {
                    "Nivel/Perfil": perfil,
                    "PDF": pdf_cnt,
                    "Excel": excel_cnt,
                    "Estado": estado,
                }
            )

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

    rec_result = payroll_obj._reconcile_data(df_despr, df_trans, df_seg)
    if isinstance(rec_result, tuple) and len(rec_result) >= 2:
        df_t, df_s = rec_result[0], rec_result[1]
    else:
        df_t, df_s = rec_result, pd.DataFrame()

    if mode == "transferencias":
        df_display = df_t.copy() if df_t is not None else pd.DataFrame()
    else:
        df_display = df_s.copy() if df_s is not None else pd.DataFrame()

    if df_display is None or df_display.empty:
        return pd.DataFrame(), None

    def normalize_list_like(v):
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
            return " ".join(parts)

        try:
            if _np is not None and isinstance(v, _np.generic):
                return str(int(v))
            if isinstance(v, (int, float)):
                return str(int(v))
        except Exception:
            pass

        return str(v)

    list_cols = [c for c in ["Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"]]
    df_tmp = df_display.copy()

    for col in [c for c in list_cols if c in df_tmp.columns]:
        def fmt_cell(x):
            s = normalize_list_like(x)
            try:
                return payroll_obj.formatear_valores(s)
            except Exception:
                return s

        df_tmp[col] = df_tmp[col].apply(fmt_cell)

    return df_tmp, None


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
                    st.markdown(build_colored_table(df_display), unsafe_allow_html=True)

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
                    df_display, error = build_reconciliation_rows(
                        str(dir_despr),
                        str(dir_trans),
                        str(dir_seg),
                        "transferencias",
                    )
                    if error:
                        st.warning(error)
                    elif df_display is None or df_display.empty:
                        st.info("No se encontraron registros o diferencias.")
                    else:
                        st.markdown(build_colored_table(df_display), unsafe_allow_html=True)
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
                    df_display, error = build_reconciliation_rows(
                        str(dir_despr),
                        str(dir_trans),
                        str(dir_seg),
                        "seguridad",
                    )
                    if error:
                        st.warning(error)
                    elif df_display is None or df_display.empty:
                        st.info("No se encontraron registros o diferencias.")
                    else:
                        st.markdown(build_colored_table(df_display), unsafe_allow_html=True)
                except Exception as exc:
                    st.error(f"Procesamiento fallido: {exc}")