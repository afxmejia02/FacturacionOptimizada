"""Processing orchestration for the web UI.

This module bridges the Streamlit front end (``app.py``) with the existing
extraction logic that lives in ``facturacion/gui_validation_app.py`` (PDF/Excel
validation) and ``mapa-de-cargos/gui_app.py`` (payroll reconciliation). It owns
the temp-file handling and the cross-PDF/Excel reconciliation, but contains no
Streamlit calls so it stays testable.
"""
from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd

from rendering import build_colored_table, build_mano_obra_table, format_count, format_dataframe

# Ensure the parent workspace is importable so we can reuse the desktop modules.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from facturacion import gui_validation_app as validator  # noqa: E402
from codigos import excluded_codes  # noqa: E402

DEBUG_MODE = os.environ.get("VALIDATION_DEBUG", "1") == "1"


def _debug_print(message: str) -> None:
    if DEBUG_MODE:
        print(f"[DEBUG][web_ui] {message}")


def _new_validator():
    """Instantiate the validator without running its tkinter ``__init__``."""
    return validator.ServicesValidationApp.__new__(validator.ServicesValidationApp)


def load_payroll_module():
    """Dynamically load ``mapa-de-cargos/gui_app.py`` (its folder name isn't importable)."""
    payroll_path = ROOT / "mapa-de-cargos" / "gui_app.py"
    if not payroll_path.exists():
        raise FileNotFoundError(
            "No se encontró el módulo de conciliación (mapa-de-cargos/gui_app.py)."
        )

    spec = importlib.util.spec_from_file_location("payroll_module", str(payroll_path))
    payroll = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(payroll)
    return payroll


def load_mano_obra_module():
    """Dynamically load ``mapa-de-cargos/mano_obra.py`` (its folder isn't importable)."""
    module_path = ROOT / "mapa-de-cargos" / "mano_obra.py"
    if not module_path.exists():
        raise FileNotFoundError(
            "No se encontró el módulo de mano de obra (mapa-de-cargos/mano_obra.py)."
        )

    spec = importlib.util.spec_from_file_location("mano_obra_module", str(module_path))
    mano_obra = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mano_obra)
    return mano_obra


# ---------------------------------------------------------------------------
# Perfiles (PDF vs Excel by date)
# ---------------------------------------------------------------------------

def _extract_perfiles_by_date(pdf_path: str) -> pd.DataFrame:
    registros = []
    excluded_profiles = {"none", "observaciones"}
    validator_obj = _new_validator()

    with tempfile.TemporaryDirectory(prefix="web_ui_perfiles_") as _tmp_dir:
        with open(pdf_path, "rb") as pdf_handle:
            pdf_bytes = pdf_handle.read()

        tmp_pdf_path = os.path.join(_tmp_dir, "upload.pdf")
        with open(tmp_pdf_path, "wb") as temp_pdf:
            temp_pdf.write(pdf_bytes)

        import pdfplumber

        with pdfplumber.open(tmp_pdf_path) as pdf:
            for page in pdf.pages:
                for tabla in page.extract_tables() or []:
                    if not tabla or len(tabla) <= 7:
                        continue

                    header = tabla[6]
                    header_norm = [
                        validator_obj._normalizar_busqueda(celda).replace(" ", "") if celda else ""
                        for celda in header
                    ]
                    if "nivel/perfil" not in header_norm:
                        continue

                    idx_perfil = header_norm.index("nivel/perfil")
                    fecha_detectada = validator_obj._normalizar_fecha(header[-1]) if header else None
                    if fecha_detectada is None:
                        continue

                    for row in tabla[7:]:
                        if len(row) <= idx_perfil:
                            continue
                        if row[4] in excluded_codes:
                            continue

                        perfil = row[idx_perfil]
                        observacion = row[-1]
                        perfil_norm = None

                        try:
                            tabla_info = str(tabla[4][2])
                        except Exception:
                            tabla_info = ""

                        tabla_info_upper = tabla_info.upper()
                        if "GLOBAL" in tabla_info_upper or "NO FACTURABLE" in tabla_info_upper:
                            continue

                        if isinstance(perfil, str) and observacion == "":
                            perfil = perfil.strip()
                            if perfil:
                                perfil_norm = validator_obj._normalizar_perfil(perfil)
                        elif observacion != "":
                            perfil = str(observacion).split()[-1]
                            if perfil:
                                perfil_norm = validator_obj._normalizar_perfil(perfil)

                        if not perfil_norm:
                            continue

                        cantidad = 1 / 3 if "24" in tabla_info else 1
                        registros.append(
                            {
                                "FECHA": fecha_detectada,
                                "PERFIL_NORM": perfil_norm,
                                "Nivel/Perfil": perfil_norm,
                                "PDF": cantidad,
                            }
                        )

    if not registros:
        _debug_print("No se extrajeron registros de perfiles por fecha desde el PDF.")
        return pd.DataFrame(columns=["FECHA", "PERFIL_NORM", "Nivel/Perfil", "PDF"])

    df = pd.DataFrame(registros)
    _debug_print(f"Registros de perfiles por fecha extraidos: {len(df)}")
    return df.groupby(["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False)["PDF"].sum()


def _extract_excel_perfiles_by_date(excel_path: str) -> pd.DataFrame:
    validator_obj = _new_validator()
    excluded_profiles = {"none", "observaciones"}
    df_hist = validator_obj._leer_excel_facturacion(excel_path)
    if "DESCRIPCION TARIFA" not in df_hist.columns:
        raise KeyError("El archivo Excel no contiene la columna 'DESCRIPCION TARIFA'.")

    df_niveles = df_hist[df_hist["DESCRIPCION TARIFA"].notna()].copy()
    df_niveles = df_niveles[
        df_niveles["DESCRIPCION TARIFA"].astype(str).str.contains("Nivel|Perfil", na=False)
    ].copy()

    import datetime as _dt

    cols_fecha = [
        col for col in df_niveles.columns if isinstance(col, (pd.Timestamp, _dt.datetime))
    ]
    if not cols_fecha:
        raise ValueError("No se detectaron columnas de fecha en el archivo Excel.")

    cols_id = [col for col in df_niveles.columns if col not in cols_fecha]
    df_largo = df_niveles.melt(
        id_vars=cols_id, value_vars=cols_fecha, var_name="FECHA", value_name="VALOR"
    )
    df_largo["FECHA"] = pd.to_datetime(df_largo["FECHA"], errors="coerce").dt.normalize()
    df_largo = df_largo[df_largo["VALOR"].notna()].copy()
    df_largo = df_largo[df_largo["VALOR"] != 0].copy()
    df_largo["PERFIL_NORM"] = df_largo["DESCRIPCION TARIFA"].apply(validator_obj._normalizar_perfil)
    df_largo = df_largo[
        ~df_largo["PERFIL_NORM"].astype(str).str.strip().str.lower().isin(excluded_profiles)
    ].copy()
    df_largo["Nivel/Perfil"] = df_largo["PERFIL_NORM"]
    df_largo = df_largo.groupby(
        ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
    )["VALOR"].sum()
    return df_largo.rename(columns={"VALOR": "Excel"})


def build_perfiles_table(pdf_sources, excel_source) -> pd.DataFrame:
    """Cross PDF profile counts against the Excel history, grouped by date.

    ``pdf_sources`` puede ser un único PDF o una lista; los conteos de todos se
    acumulan antes de cruzar contra el Excel.
    """
    if not isinstance(pdf_sources, (list, tuple)):
        pdf_sources = [pdf_sources]

    empty = pd.DataFrame(columns=["Fecha", "Nivel/Perfil", "PDF", "Excel", "Estado"])
    with tempfile.TemporaryDirectory(prefix="web_ui_perfiles_table_") as tmp_dir:
        excel_path = os.path.join(tmp_dir, "upload.xlsx")

        if hasattr(excel_source, "getbuffer"):
            with open(excel_path, "wb") as excel_handle:
                excel_handle.write(excel_source.getbuffer())
        else:
            excel_path = str(excel_source)

        partes_pdf = []
        for idx, pdf_source in enumerate(pdf_sources):
            if hasattr(pdf_source, "getbuffer"):
                pdf_path = os.path.join(tmp_dir, f"upload_{idx}.pdf")
                with open(pdf_path, "wb") as pdf_handle:
                    pdf_handle.write(pdf_source.getbuffer())
            else:
                pdf_path = str(pdf_source)
            df_parte = _extract_perfiles_by_date(pdf_path)
            if not df_parte.empty:
                partes_pdf.append(df_parte)

        df_excel = _extract_excel_perfiles_by_date(excel_path)

        if not partes_pdf or df_excel.empty:
            return empty

        df_pdf = pd.concat(partes_pdf, ignore_index=True).groupby(
            ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
        )["PDF"].sum()

        df_merge = df_pdf.merge(
            df_excel, on=["FECHA", "PERFIL_NORM", "Nivel/Perfil"], how="outer"
        )
        df_merge["PDF"] = df_merge["PDF"].fillna(0)
        df_merge["Excel"] = df_merge["Excel"].fillna(0)
        df_merge["Estado"] = df_merge.apply(
            lambda row: "OK" if row["PDF"] == row["Excel"] else "Valores diferentes", axis=1
        )
        df_merge = df_merge.rename(columns={"FECHA": "Fecha"})
        df_merge["Fecha"] = pd.to_datetime(df_merge["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d")
        df_merge = df_merge[
            ~df_merge["Nivel/Perfil"].astype(str).str.strip().str.lower().isin(
                {"none", "incapacidad", "observaciones"}
            )
        ].copy()
        df_merge = df_merge[["Fecha", "Nivel/Perfil", "PDF", "Excel", "Estado"]]
        return df_merge.sort_values(["Fecha", "Nivel/Perfil"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Pagos (equipos / servicios / perfiles)
# ---------------------------------------------------------------------------

def process_pagos(pdf_files, excel_file, tipo):
    """Compare PDF counts against the Excel history for equipos/servicios/perfiles.

    ``pdf_files`` puede ser un único archivo o una lista de varios PDFs; los
    conteos de todos ellos se acumulan antes de cruzar contra el Excel.

    Returns ``(df_display, message)``: a results DataFrame (with a ``Fecha`` and
    ``Estado`` column, ready for the date filter + coloured table) or, when there
    is nothing to show, ``message`` explains why and ``df_display`` is ``None``.
    """
    if not isinstance(pdf_files, (list, tuple)):
        pdf_files = [pdf_files]

    _debug_print(
        f"Inicio process_pagos. tipo={tipo}, "
        f"pdfs={[getattr(f, 'name', 'N/A') for f in pdf_files]}, "
        f"excel={getattr(excel_file, 'name', 'N/A')}"
    )
    validator_obj = _new_validator()

    with tempfile.TemporaryDirectory(prefix="web_ui_") as tmp_dir:
        excel_path = os.path.join(tmp_dir, "upload.xlsx")
        with open(excel_path, "wb") as excel_handle:
            excel_handle.write(excel_file.getbuffer())

        pdf_paths = []
        for idx, pdf_file in enumerate(pdf_files):
            pdf_path = os.path.join(tmp_dir, f"upload_{idx}.pdf")
            with open(pdf_path, "wb") as pdf_handle:
                pdf_handle.write(pdf_file.getbuffer())
            pdf_paths.append(pdf_path)

        rows = []

        if tipo == "perfiles":
            df_display = build_perfiles_table(pdf_paths, excel_path)
            _debug_print(f"Tabla perfiles construida. filas={len(df_display)}")
            if df_display.empty:
                return None, "No se encontraron perfiles en el PDF o no fue posible cruzarlos por fecha."
            return df_display, None
        else:
            partes_pdf = []
            for pdf_path in pdf_paths:
                df_parte = validator_obj._extraer_conteo_pdf(pdf_path, tipo)
                if df_parte is not None and not df_parte.empty:
                    partes_pdf.append(df_parte)
            if not partes_pdf:
                return None, f"No se encontraron elementos válidos de tipo {tipo} en el PDF."

            df_pdf = pd.concat(partes_pdf, ignore_index=True)
            pdf_agg = df_pdf.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

            for fecha in sorted(pdf_agg["FECHA"].dropna().unique()):
                conteo_excel = validator_obj._extraer_conteo_excel(excel_path, fecha)
                pdf_fecha = pdf_agg[pdf_agg["FECHA"] == fecha]
                all_servicios = sorted(
                    pdf_fecha["TIPO DE EQUIPO"].dropna().astype(str).unique().tolist()
                )

                for servicio in all_servicios:
                    pdf_match = pdf_fecha[pdf_fecha["TIPO DE EQUIPO"] == servicio]
                    pdf_cnt = format_count(pdf_match["CANTIDAD"].sum()) if not pdf_match.empty else 0
                    if float(pdf_cnt or 0) == 0:
                        continue
                    # Emparejar por clave robusta (tolera comillas/conjunciones).
                    excel_cnt = format_count(
                        conteo_excel.get(validator_obj._clave_equipo(servicio), 0)
                    )
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
            _debug_print("No hay datos para comparar en process_pagos.")
            return None, "No se encontraron datos para comparar."

        _debug_print(f"Tabla final construida en process_pagos. filas={len(df_display)}")
        return df_display, None


# ---------------------------------------------------------------------------
# Reconciliation (mapa de cargos: transferencias / seguridad social)
# ---------------------------------------------------------------------------

def _save_uploads(uploaded_files, target_dir):
    for uploaded_file in uploaded_files:
        if uploaded_file and uploaded_file.name:
            with open(os.path.join(target_dir, uploaded_file.name), "wb") as handle:
                handle.write(uploaded_file.getbuffer())


def process_reconciliation(despr_files, trans_files, seguridad_files, recon_mode, transfer_format="tabarca"):
    """Reconcile payslips against transfers or social-security (IBC).

    Returns ``(parts, message, df_transfers, df_seguridad, tables)`` where
    ``parts`` is a list of ``(title, html_table)`` for display and ``tables`` is
    the matching list of ``(title, DataFrame)`` used to build the PDF export.
    """
    payroll = load_payroll_module()
    PayrollApp = payroll.PayrollReconciliationApp
    payroll_obj = PayrollApp.__new__(PayrollApp)

    with tempfile.TemporaryDirectory(prefix="web_ui_rec_") as tmp_base:
        dir_despr = os.path.join(tmp_base, "despr")
        dir_trans = os.path.join(tmp_base, "trans")
        dir_seg = os.path.join(tmp_base, "seg")
        os.makedirs(dir_despr, exist_ok=True)
        os.makedirs(dir_trans, exist_ok=True)
        os.makedirs(dir_seg, exist_ok=True)

        _save_uploads(despr_files, dir_despr)
        if recon_mode == "transfers":
            _save_uploads(trans_files, dir_trans)
        else:
            _save_uploads(seguridad_files, dir_seg)

        # The desprendibles layout must match the transfer layout (TABARCA / ITALCO).
        df_despr = payroll_obj._process_desprendibles(dir_despr, transfer_format)
        df_trans = (
            payroll_obj._process_transferencia(dir_trans, transfer_format)
            if os.listdir(dir_trans)
            else None
        )
        df_seg = (
            payroll_obj.procesar_seguridad_social(dir_seg, transfer_format)
            if os.listdir(dir_seg)
            else None
        )

        reconcile_result = payroll_obj._reconcile_data(df_despr, df_trans, df_seg)
        if isinstance(reconcile_result, tuple) and len(reconcile_result) == 2:
            df_transfers, df_seguridad = reconcile_result
        else:
            df_transfers = pd.DataFrame()
            df_seguridad = pd.DataFrame()

        if (df_transfers is None or df_transfers.empty) and (df_seguridad is None or df_seguridad.empty):
            return [], "No se encontraron registros o diferencias.", df_transfers, df_seguridad, []

        df_display_t = df_transfers.copy() if df_transfers is not None else pd.DataFrame()
        df_display_s = df_seguridad.copy() if df_seguridad is not None else pd.DataFrame()

        df_display_t = format_dataframe(df_display_t, payroll_obj.formatear_valores)
        df_display_s = format_dataframe(df_display_s, payroll_obj.formatear_valores)

        parts = []
        tables = []
        if recon_mode == "transfers" and df_display_t is not None and not df_display_t.empty:
            parts.append(("Revisión Transferencias", build_colored_table(df_display_t)))
            tables.append(("Revisión Transferencias", df_display_t))
        if recon_mode == "seguridad" and df_display_s is not None and not df_display_s.empty:
            parts.append(("Revisión Seguridad Social (IBC)", build_colored_table(df_display_s)))
            tables.append(("Revisión Seguridad Social (IBC)", df_display_s))

        if not parts:
            return [], "No se encontraron registros para el modo seleccionado.", df_transfers, df_seguridad, []

        return parts, None, df_transfers, df_seguridad, tables


# ---------------------------------------------------------------------------
# Mano de obra (mapa de cargos: Informe de Costo vs registro de la ODS)
# ---------------------------------------------------------------------------

def process_mano_obra(informe_file, ods_file):
    """Cross the Informe de Costo against the ODS registry, per worker.

    Returns ``(table_html, message, df_result)`` where ``df_result`` keeps the
    list-valued comparison cells (one element = match, two = mismatch) so the
    Excel export can colour the same cells the HTML table highlights.
    """
    mano_obra = load_mano_obra_module()

    with tempfile.TemporaryDirectory(prefix="web_ui_mano_obra_") as tmp_dir:
        informe_path = os.path.join(tmp_dir, "informe.xlsx")
        ods_path = os.path.join(tmp_dir, "ods.xlsx")
        with open(informe_path, "wb") as handle:
            handle.write(informe_file.getbuffer())
        with open(ods_path, "wb") as handle:
            handle.write(ods_file.getbuffer())

        df_result = mano_obra.comparar_mano_obra(informe_path, ods_path)

    if df_result is None or df_result.empty:
        _debug_print("Mano de obra: ninguna persona cruzada entre Informe y ODS.")
        return None, "Ninguna persona del Informe se encontró en la ODS (cruce por documento).", df_result

    table_html = build_mano_obra_table(df_result)
    _debug_print(f"Tabla mano de obra construida. filas={len(df_result)}")
    return table_html, None, df_result
