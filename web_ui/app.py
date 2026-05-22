"""Simple Flask web UI that reuses functions from `facturacion/gui_validation_app.py`.

This app accepts a PDF and an Excel file, a selection for `tipo` (perfiles/equipos/servicios),
and calls the corresponding extraction/comparison functions without instantiating the GUI.
"""
from pathlib import Path
import sys
import tempfile
import os
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd

# Ensure parent workspace is on sys.path so we can import the existing module
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from facturacion import gui_validation_app as validator
import importlib.util
from pathlib import Path as _Path

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret")


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/validate", methods=["POST"])
def validate():
    app_choice = request.form.get("app_choice", "validation")
    tipo = request.form.get("tipo", "equipos")
    pdf_file = request.files.get("pdf")
    excel_file = request.files.get("excel")

    def build_colored_table(df_display: pd.DataFrame) -> str:
        headers = list(df_display.columns)
        parts = []
        parts.append('<table class="table table-striped table-bordered">')
        parts.append('<thead><tr>')
        for h in headers:
            parts.append(f'<th>{h}</th>')
        parts.append('</tr></thead>')
        parts.append('<tbody>')

        for _, row in df_display.iterrows():
            estado = str(row.get("Estado", "")).upper()
            row_class = "table-success" if estado == "OK" else "table-danger"
            parts.append(f'<tr class="{row_class}">')
            for h in headers:
                val = row[h] if pd.notna(row[h]) else ""
                parts.append(f'<td>{val}</td>')
            parts.append('</tr>')

        parts.append('</tbody></table>')
        return "".join(parts)

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
    # Branch handling: validation (facturacion) or reconciliation (mapa-de-cargos)
    if app_choice == "validation":
        if not pdf_file or pdf_file.filename == "":
            flash("Por favor sube un archivo PDF.")
            return redirect(url_for("index"))
        if not excel_file or excel_file.filename == "":
            flash("Por favor sube un archivo Excel.")
            return redirect(url_for("index"))

        tmp_dir = tempfile.mkdtemp(prefix="web_ui_")
        pdf_path = os.path.join(tmp_dir, "upload.pdf")
        excel_path = os.path.join(tmp_dir, "upload.xlsx")
        pdf_file.save(pdf_path)
        excel_file.save(excel_path)

        # Create an instance without running the tkinter __init__
        validator_obj = validator.ServicesValidationApp.__new__(validator.ServicesValidationApp)

        try:
            rows = []

            if tipo == "perfiles":
                conteo_pdf, fecha = validator_obj._extraer_perfiles_pdf(pdf_path)
                if not conteo_pdf:
                    flash("No se encontraron perfiles en el PDF.")
                    return redirect(url_for("index"))

                conteo_excel = validator_obj._extraer_conteo_excel_perfiles(excel_path, fecha)
                excluded_profiles = {"none", "incapacidad", "observaciones"}
                # Only trust what appears in the PDF; ignore Excel-only rows to avoid noise.
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
                    rows.append({
                        "Nivel/Perfil": perfil,
                        "PDF": pdf_cnt,
                        "Excel": excel_cnt,
                        "Estado": estado,
                    })

            else:
                df_pdf = validator_obj._extraer_conteo_pdf(pdf_path, tipo)
                if df_pdf is None or df_pdf.empty:
                    flash(f"No se encontraron elementos válidos de tipo {tipo} en el PDF.")
                    return redirect(url_for("index"))

                pdf_agg = df_pdf.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

                for fecha in sorted(pdf_agg["FECHA"].dropna().unique()):
                    conteo_excel = validator_obj._extraer_conteo_excel(excel_path, fecha)
                    pdf_fecha = pdf_agg[pdf_agg["FECHA"] == fecha]
                    # Only keep services that actually appear in the PDF.
                    all_servicios = sorted(pdf_fecha["TIPO DE EQUIPO"].dropna().astype(str).unique().tolist())

                    for servicio in all_servicios:
                        pdf_match = pdf_fecha[pdf_fecha["TIPO DE EQUIPO"] == servicio]
                        pdf_cnt = format_count(pdf_match["CANTIDAD"].sum()) if not pdf_match.empty else 0
                        if float(pdf_cnt or 0) == 0:
                            continue
                        excel_cnt = format_count(conteo_excel.get(servicio, 0))
                        estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
                        rows.append({
                            "Fecha": fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else str(fecha),
                            "Servicio": servicio,
                            "PDF": pdf_cnt,
                            "Excel": excel_cnt,
                            "Estado": estado,
                        })

            df_display = pd.DataFrame(rows)
            if df_display.empty:
                return render_template("results.html", table_html=None, message="No se encontraron datos para comparar.")

            table_html = build_colored_table(df_display)
            return render_template("results.html", table_html=table_html, message=None)
        except Exception as exc:
            flash(f"Procesamiento fallido: {exc}")
            return redirect(url_for("index"))
        finally:
            try:
                os.remove(pdf_path)
                os.remove(excel_path)
                os.rmdir(tmp_dir)
            except Exception:
                pass

    # Reconciliation path: accept multiple uploaded PDFs for desprendibles, transferencias, seguridad
    if app_choice == "reconciliation":
        despr_files = request.files.getlist("desprendibles")
        trans_files = request.files.getlist("transferencias")
        seguridad_files = request.files.getlist("seguridad")

        if not despr_files or all(f.filename == "" for f in despr_files):
            flash("Por favor sube al menos un PDF de desprendibles.")
            return redirect(url_for("index"))
        if not trans_files or all(f.filename == "" for f in trans_files):
            flash("Por favor sube al menos un PDF de transferencias.")
            return redirect(url_for("index"))

        tmp_base = tempfile.mkdtemp(prefix="web_ui_rec_")
        dir_despr = os.path.join(tmp_base, "despr")
        dir_trans = os.path.join(tmp_base, "trans")
        dir_seg = os.path.join(tmp_base, "seg")
        os.makedirs(dir_despr, exist_ok=True)
        os.makedirs(dir_trans, exist_ok=True)
        os.makedirs(dir_seg, exist_ok=True)

        for i, f in enumerate(despr_files):
            if f and f.filename:
                f.save(os.path.join(dir_despr, f.filename))
        for i, f in enumerate(trans_files):
            if f and f.filename:
                f.save(os.path.join(dir_trans, f.filename))
        for i, f in enumerate(seguridad_files):
            if f and f.filename:
                f.save(os.path.join(dir_seg, f.filename))

        # Dynamically load the payroll module (folder name contains hyphen; import by file)
        payroll_path = _Path(ROOT) / "mapa-de-cargos" / "gui_app.py"
        if not payroll_path.exists():
            flash("No se encontró el módulo de conciliación (mapa-de-cargos/gui_app.py).")
            return redirect(url_for("index"))

        spec = importlib.util.spec_from_file_location("payroll_module", str(payroll_path))
        payroll = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(payroll)

        # Provide a clear alias similar to 'from X import Y as Z'
        PayrollApp = payroll.PayrollReconciliationApp
        payroll_obj = PayrollApp.__new__(PayrollApp)
        

        try:
            df_despr = payroll_obj._process_desprendibles(dir_despr)
            df_trans = payroll_obj._process_transferencia(dir_trans)
            df_seg = payroll_obj.procesar_seguridad_social(dir_seg) if os.listdir(dir_seg) else None

            df = payroll_obj._reconcile_data(df_despr, df_trans, df_seg)
            if df is None or df.empty:
                return render_template("results.html", table_html=None, message="No se encontraron registros o diferencias.")

            # Prepare display DataFrame and format monetary/list values using payroll formatter
            df_display = df.copy()

            def normalize_list_like(v):
                # Convert numpy types and lists into a space-separated string of ints
                if v is None:
                    return ""
                # If it's a pandas/NumPy scalar
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

                # single numeric
                try:
                    if _np is not None and isinstance(v, _np.generic):
                        return str(int(v))
                    if isinstance(v, (int, float)):
                        return str(int(v))
                except Exception:
                    pass

                return str(v)

            # Columns that contain list-like monetary values
            list_cols = [c for c in ["Neto_desprendibles", "Valores_transferencia", "Devengado", "IBC"] if c in df_display.columns]
            for col in list_cols:
                # Use the payroll object's formatter to render COP format
                def fmt_cell(x):
                    s = normalize_list_like(x)
                    # payroll.formatear_valores expects a string or None
                    try:
                        return payroll_obj.formatear_valores(s)
                    except Exception:
                        return s
                df_display[col] = df_display[col].apply(fmt_cell)

            # Build HTML table with Bootstrap row classes based on Estado
            try:
                table_html = build_colored_table(df_display)
            except Exception:
                table_html = df_display.to_html(index=False, classes="table table-striped", escape=False)

            return render_template("results.html", table_html=table_html, message=None)
        except Exception as exc:
            flash(f"Procesamiento de conciliación falló: {exc}")
            return redirect(url_for("index"))
        finally:
            # cleanup
            try:
                for root_dir, dirs, files in os.walk(tmp_base, topdown=False):
                    for name in files:
                        os.remove(os.path.join(root_dir, name))
                    for name in dirs:
                        os.rmdir(os.path.join(root_dir, name))
                os.rmdir(tmp_base)
            except Exception:
                pass

    flash("Opción desconocida.")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, port=8500)
