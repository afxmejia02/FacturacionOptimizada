"""Orquestacion entre la UI de Streamlit (``app.py``) y los paquetes de logica.

Usa ``facturacion`` (PDF/Excel) y ``nomina`` (conciliacion y mano de obra). Se
encarga de los archivos temporales y del cruce; no llama a Streamlit, para poder
probarse aparte.
"""
from __future__ import annotations

import os
import re
import sys
import tempfile
import traceback
from pathlib import Path

import pandas as pd

from rendering import build_colored_table, build_mano_obra_table, format_count, format_dataframe

# La raiz del repo debe estar en sys.path para importar facturacion/ y nomina/.
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from facturacion import (  # noqa: E402
    clave_equipo,
    col_codigo_tarifa,
    es_celda_vacia,
    extraer_conteo_pdf,
    leer_excel_facturacion,
    leer_histograma_largo,
    normalizar_busqueda,
    normalizar_fecha,
    normalizar_perfil,
    parsear_observacion_perfil,
    prefijos_seccion_pdf,
)
from codigos import excluded_codes  # noqa: E402
from nomina import (  # noqa: E402
    conciliar,
    formatear_valores,
    procesar_desprendibles,
    procesar_seguridad_social,
    procesar_transferencias,
)
from nomina import mano_obra  # noqa: E402

DEBUG_MODE = os.environ.get("VALIDATION_DEBUG", "1") == "1"


def _debug_print(message: str) -> None:
    if DEBUG_MODE:
        print(f"[DEBUG][web_ui] {message}")


# ---------------------------------------------------------------------------
# Perfiles (PDF vs Excel by date)
# ---------------------------------------------------------------------------

#: Filas del encabezado de la planilla: "Nivel/ Perfil" y la fecha van en la 6,
#: y "CEDULA | TRABAJADOR | CARGO" en la 7. Los datos empiezan en la 8.
_FILAS_ENCABEZADO_PLANILLA = (6, 7)


def _indice_columna_cargo(tabla) -> int | None:
    """Indice de la columna CARGO de la planilla, o ``None`` si no aparece.

    Solo se miran las filas del encabezado: mas abajo esta el "CARGO:" del
    responsable de la orden de servicio, que es otra cosa.
    """
    for fila in _FILAS_ENCABEZADO_PLANILLA:
        if fila >= len(tabla):
            continue
        for i, celda in enumerate(tabla[fila]):
            if celda and normalizar_busqueda(celda).replace(" ", "") == "cargo":
                return i
    return None


def _extract_perfiles_by_date(pdf_path: str) -> pd.DataFrame:
    registros = []
    excluded_profiles = {"none", "observaciones"}

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
                        normalizar_busqueda(celda).replace(" ", "") if celda else ""
                        for celda in header
                    ]
                    if "nivel/perfil" not in header_norm:
                        continue

                    idx_perfil = header_norm.index("nivel/perfil")
                    # Los perfiles con nombre largo no caben en "Nivel/Perfil" y
                    # el nombre real queda en "Cargo" (ver _aplicar_respaldo_cargo).
                    # El encabezado ocupa DOS filas: "Nivel/ Perfil" esta en la 6
                    # y "CEDULA | TRABAJADOR | CARGO" en la 7, asi que hay que
                    # mirar ambas. No se busca mas abajo para no confundirlo con
                    # el "CARGO:" del responsable, que va en la cabecera del acta.
                    idx_cargo = _indice_columna_cargo(tabla)
                    fecha_detectada = normalizar_fecha(header[-1]) if header else None
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

                        # Interpretar la columna Observaciones (recategorización,
                        # "E y F", "NO FACTURABLE" y "24h", que pueden coexistir).
                        recategorizado, es_ef, no_facturable, es_24h_obs = (
                            parsear_observacion_perfil(observacion)
                        )
                        if no_facturable:
                            continue  # la observación indica que no se factura

                        if recategorizado:
                            fuente = recategorizado
                        elif es_ef or es_24h_obs or es_celda_vacia(observacion):
                            # "E y F", "24 horas" o sin observación (celda vacía,
                            # incluida None/espacios): el nivel es el de la columna,
                            # no la última palabra de la observación.
                            fuente = perfil.strip() if isinstance(perfil, str) else perfil
                        else:
                            # Otra observación no reconocida: comportamiento previo.
                            fuente = str(observacion).split()[-1]
                        if fuente:
                            perfil_norm = normalizar_perfil(fuente)

                        if not perfil_norm:
                            continue

                        # 24 horas (al inicio de la hoja o en observaciones): 1/3 por
                        # persona, salvo el marcador "E y F" (cuenta como 1 unidad).
                        es_24h = ("24" in tabla_info) or es_24h_obs
                        cantidad = 1 / 3 if (es_24h and not es_ef) else 1
                        cargo = row[idx_cargo] if idx_cargo is not None and len(row) > idx_cargo else None
                        registros.append(
                            {
                                "FECHA": fecha_detectada,
                                "PERFIL_NORM": perfil_norm,
                                "Nivel/Perfil": perfil_norm,
                                "CARGO_NORM": normalizar_perfil(cargo) if cargo else None,
                                "PDF": cantidad,
                            }
                        )

    if not registros:
        _debug_print("No se extrajeron registros de perfiles por fecha desde el PDF.")
        return pd.DataFrame(
            columns=["FECHA", "PERFIL_NORM", "Nivel/Perfil", "CARGO_NORM", "PDF"]
        )

    df = pd.DataFrame(registros)
    _debug_print(f"Registros de perfiles por fecha extraidos: {len(df)}")
    # El cargo se conserva por grupo (first) para poder usarlo como respaldo.
    return df.groupby(
        ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False, dropna=False
    ).agg(PDF=("PDF", "sum"), CARGO_NORM=("CARGO_NORM", "first"))


def _extract_excel_perfiles_by_date(excel_path: str) -> pd.DataFrame:
    excluded_profiles = {"none", "observaciones"}
    df_hist = leer_excel_facturacion(excel_path)
    if "DESCRIPCION TARIFA" not in df_hist.columns:
        raise KeyError("El archivo Excel no contiene la columna 'DESCRIPCION TARIFA'.")

    df_niveles = df_hist[df_hist["DESCRIPCION TARIFA"].notna()].copy()
    # Perfiles = mano de obra (COD. TAR. 5.1 y 5.4). Se exige un código de tarifa
    # real que empiece por "5." y se excluyen equipos (5.5) y servicios (5.6).
    # Exigir el código descarta filas sin tarifa (encabezados, firmas de pie de
    # página como "Vo.Bo. Lider…" o nombres) e incluye tarifas de mano de obra que
    # no dicen "Nivel/Perfil" (p. ej. "Inspector certificado: API/ASME NACIONAL").
    col_cod = col_codigo_tarifa(df_niveles)
    if col_cod is not None:
        cods = df_niveles[col_cod].astype(str).str.strip()
        es_perfil = cods.str.startswith("5.") & ~(
            cods.str.startswith("5.5") | cods.str.startswith("5.6")
        )
        df_niveles = df_niveles[es_perfil].copy()
    else:
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
    df_largo["VALOR"] = pd.to_numeric(df_largo["VALOR"], errors="coerce")
    df_largo = df_largo[df_largo["VALOR"].notna()].copy()
    df_largo = df_largo[df_largo["VALOR"] != 0].copy()
    df_largo["PERFIL_NORM"] = df_largo["DESCRIPCION TARIFA"].apply(normalizar_perfil)
    df_largo = df_largo[
        ~df_largo["PERFIL_NORM"].astype(str).str.strip().str.lower().isin(excluded_profiles)
    ].copy()
    df_largo["Nivel/Perfil"] = df_largo["PERFIL_NORM"]
    df_largo = df_largo.groupby(
        ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
    )["VALOR"].sum()
    return df_largo.rename(columns={"VALOR": "Excel"})


def _clave_perfil(valor) -> str:
    """Forma canonica para decidir si dos nombres de perfil son el mismo."""
    return re.sub(r"\s+", " ", str(valor)).strip().casefold()


def _aplicar_respaldo_cargo(df_pdf: pd.DataFrame, df_excel: pd.DataFrame) -> pd.DataFrame:
    """Reasigna al nombre del Cargo los perfiles del PDF que no estan en el Excel.

    Un perfil de nombre largo (``Inspector certificado: API/ASME NACIONAL``) no
    cabe en la columna ``Nivel/Perfil`` del PDF y llega partido
    (``Inspector certificad o:``), mientras el nombre completo si esta en
    ``Cargo``. Sin esto aparecen dos filas descuadradas: el nombre partido con
    PDF=1/Excel=0 y el completo con PDF=0/Excel=1.

    Solo se toca lo que **no** cruza por ``Nivel/Perfil`` y cuyo Cargo coincide
    exactamente (ignorando mayusculas y espacios repetidos) con un nombre del
    Excel, asi que un perfil que ya cruzaba no cambia.
    """
    if df_pdf.empty or "CARGO_NORM" not in df_pdf.columns or df_excel.empty:
        return df_pdf

    del_excel = {_clave_perfil(v): v for v in df_excel["PERFIL_NORM"]}
    del_pdf = {_clave_perfil(v) for v in df_pdf["PERFIL_NORM"]}

    df_pdf = df_pdf.copy()
    for pos, fila in df_pdf.iterrows():
        if _clave_perfil(fila["PERFIL_NORM"]) in del_excel:
            continue  # ya cruza por Nivel/Perfil
        cargo = fila.get("CARGO_NORM")
        if not isinstance(cargo, str) or not cargo.strip():
            continue
        clave = _clave_perfil(cargo)
        if clave in del_excel and clave not in del_pdf:
            nombre = del_excel[clave]
            _debug_print(
                f"Perfil '{fila['PERFIL_NORM']}' no esta en el Excel; se cruza "
                f"por su Cargo '{nombre}'."
            )
            df_pdf.at[pos, "PERFIL_NORM"] = nombre
            df_pdf.at[pos, "Nivel/Perfil"] = nombre

    return df_pdf.groupby(
        ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
    )["PDF"].sum()


def build_perfiles_table(pdf_sources, excel_sources) -> pd.DataFrame:
    """Cross PDF profile counts against the Excel history, grouped by date.

    ``pdf_sources`` / ``excel_sources`` pueden ser un único archivo o una lista;
    los conteos de todos los PDFs y la planilla de todos los Excel se acumulan
    antes de cruzar.
    """
    if not isinstance(pdf_sources, (list, tuple)):
        pdf_sources = [pdf_sources]
    if not isinstance(excel_sources, (list, tuple)):
        excel_sources = [excel_sources]

    empty = pd.DataFrame(columns=["Fecha", "Nivel/Perfil", "PDF", "Excel", "Estado"])
    with tempfile.TemporaryDirectory(prefix="web_ui_perfiles_table_") as tmp_dir:
        excel_paths = []
        for idx, excel_source in enumerate(excel_sources):
            if hasattr(excel_source, "getbuffer"):
                excel_path = os.path.join(tmp_dir, f"upload_{idx}.xlsx")
                with open(excel_path, "wb") as excel_handle:
                    excel_handle.write(excel_source.getbuffer())
            else:
                excel_path = str(excel_source)
            excel_paths.append(excel_path)

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

        partes_excel = [_extract_excel_perfiles_by_date(p) for p in excel_paths]
        partes_excel = [df for df in partes_excel if not df.empty]

        if not partes_pdf or not partes_excel:
            return empty

        df_pdf = pd.concat(partes_pdf, ignore_index=True).groupby(
            ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False, dropna=False
        ).agg(PDF=("PDF", "sum"), CARGO_NORM=("CARGO_NORM", "first"))
        df_excel = pd.concat(partes_excel, ignore_index=True).groupby(
            ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
        )["Excel"].sum()

        df_pdf = _aplicar_respaldo_cargo(df_pdf, df_excel)

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
# Perfiles ITALCO (planilla de personal PDF vs histograma Excel por fecha)
# ---------------------------------------------------------------------------
#
# Diferencias frente a TABARCA:
#   - El nivel/perfil se toma del **Cargo** de la planilla (``PAILERO 1A E11`` →
#     ``E11``), no de una columna "Nivel/Perfil" ya normalizada.
#   - No hay recategorización por observaciones ni casos GLOBAL / 24 horas: cada
#     fila es una persona (cantidad 1).
#   - En el histograma el equivalente a ``COD. TAR.`` es ``ITEM PAGO``; al revisar
#     perfiles se excluyen los ítems de equipos (5.5) y servicios (5.6) para que
#     el histograma no cuente equipos ni servicios.

def _fecha_iso_italco(valor):
    """Fecha del reporte ITALCO (``FECHA REPORTE``) en formato ISO ``AAAA-MM-DD``.

    Se parsea como año-mes-día explícito; el parser general de TABARCA usa
    ``dayfirst=True`` y confundiría ``2026-06-08`` con el 6 de agosto.
    """
    if valor is None:
        return None
    match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(valor).strip())
    if not match:
        return None
    try:
        return pd.Timestamp(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def _codigo_perfil_italco(texto):
    """Código de nivel al final de un texto ITALCO.

    En el PDF el Cargo trae el oficio completo (``PAILERO 1A E11``, ``OBRERO A2``)
    y en el histograma la CATEGORÍA trae ``NIVEL E11``; en ambos el nivel es el
    código final (``E11``, ``A2``). Se toma solo ese código, igual que TABARCA
    toma ``E11`` / ``A2``. Devuelve ``None`` si el texto no termina en un código,
    lo que descarta filas de pie de página o categorías sin nivel.
    """
    if not isinstance(texto, str):
        return None
    match = re.search(r"([A-Z]\d{1,2})$", texto.strip().upper())
    return match.group(1) if match else None


def _extract_perfiles_italco_by_date(pdf_path: str) -> pd.DataFrame:
    registros = []

    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for tabla in page.extract_tables() or []:
                fecha = None
                idx_cargo = None
                header_idx = None
                for i, row in enumerate(tabla):
                    cells = [celda if celda is not None else "" for celda in row]
                    norm = [normalizar_busqueda(c).replace(" ", "") for c in cells]
                    if fecha is None and "fechareporte" in norm:
                        j = norm.index("fechareporte")
                        for siguiente in cells[j + 1:]:
                            if siguiente and str(siguiente).strip():
                                fecha = _fecha_iso_italco(siguiente)
                                if fecha is not None:
                                    break
                    if idx_cargo is None and "cargo" in norm:
                        idx_cargo = norm.index("cargo")
                        header_idx = i
                if idx_cargo is None or fecha is None or header_idx is None:
                    continue

                for row in tabla[header_idx + 1:]:
                    if len(row) <= idx_cargo:
                        continue
                    codigo = _codigo_perfil_italco(row[idx_cargo])
                    if not codigo:
                        continue
                    registros.append(
                        {
                            "FECHA": fecha,
                            "PERFIL_NORM": codigo,
                            "Nivel/Perfil": codigo,
                            "PDF": 1,
                        }
                    )

    if not registros:
        _debug_print("No se extrajeron perfiles ITALCO desde el PDF.")
        return pd.DataFrame(columns=["FECHA", "PERFIL_NORM", "Nivel/Perfil", "PDF"])

    df = pd.DataFrame(registros)
    _debug_print(f"Registros de perfiles ITALCO extraidos: {len(df)}")
    return df.groupby(["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False)["PDF"].sum()


def _leer_histograma_italco(excel_path: str) -> pd.DataFrame:
    """Lee el histograma ITALCO detectando la fila de encabezado real
    (``ITEM PAGO | ... | CATEGORÍA | DESCRIPCIÓN | ...``), que no es la primera."""
    crudo = pd.read_excel(excel_path, header=None)
    fila_encabezado = 0
    for i in range(min(25, len(crudo))):
        celdas = {
            normalizar_busqueda(v).replace(" ", "") for v in crudo.iloc[i].tolist()
        }
        if "itempago" in celdas or "categoria" in celdas:
            fila_encabezado = i
            break
    return pd.read_excel(excel_path, header=fila_encabezado)


def _extract_excel_perfiles_italco_by_date(excel_path: str) -> pd.DataFrame:
    import datetime as _dt

    df_hist = _leer_histograma_italco(excel_path)

    def _col(clave):
        return next(
            (c for c in df_hist.columns if clave in normalizar_busqueda(c).replace(" ", "")),
            None,
        )

    col_item = _col("itempago")
    col_categoria = _col("categoria")
    if col_categoria is None:
        raise KeyError("El histograma ITALCO no contiene la columna 'CATEGORÍA'.")

    cols_fecha = [c for c in df_hist.columns if isinstance(c, (pd.Timestamp, _dt.datetime))]
    if not cols_fecha:
        raise ValueError("No se detectaron columnas de fecha en el histograma ITALCO.")

    df = df_hist.copy()
    df["PERFIL_NORM"] = df[col_categoria].apply(_codigo_perfil_italco)
    df = df[df["PERFIL_NORM"].notna()].copy()

    # Excluir equipos (ITEM PAGO 5.5*) y servicios (5.6*) para que el histograma
    # no cuente equipos ni servicios al revisar perfiles.
    if col_item is not None:
        item = df[col_item].astype(str).str.strip()
        df = df[~(item.str.startswith("5.5") | item.str.startswith("5.6"))].copy()

    df["Nivel/Perfil"] = df["PERFIL_NORM"]

    # Ítems de 24 horas: el ITEM PAGO 5.1.1.4 corresponde a "MANO DE OBRA DIRECTA
    # 24 HR", donde una persona cubre 3 turnos, así que el histograma reparte su
    # cantidad en tercios. Se multiplica por 3 y se redondea al entero más cercano
    # (antes de sumar) para que coincida con el conteo por persona del PDF.
    if col_item is not None:
        item = df[col_item].astype(str).str.strip()
        df["_ES_24H"] = item.eq("5.1.1.4") | item.str.startswith("5.1.1.4.")
    else:
        df["_ES_24H"] = False

    df_largo = df.melt(
        id_vars=["PERFIL_NORM", "Nivel/Perfil", "_ES_24H"],
        value_vars=cols_fecha,
        var_name="FECHA",
        value_name="VALOR",
    )
    df_largo["FECHA"] = pd.to_datetime(df_largo["FECHA"], errors="coerce").dt.normalize()
    df_largo["VALOR"] = pd.to_numeric(df_largo["VALOR"], errors="coerce")
    df_largo = df_largo[df_largo["VALOR"].notna() & (df_largo["VALOR"] != 0)].copy()
    es_24h = df_largo["_ES_24H"]
    # +0.5 y truncado = redondeo al entero más cercano (valores no negativos).
    df_largo.loc[es_24h, "VALOR"] = (df_largo.loc[es_24h, "VALOR"] * 3 + 0.5).astype(int)
    df_largo = df_largo.groupby(
        ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
    )["VALOR"].sum()
    return df_largo.rename(columns={"VALOR": "Excel"})


def build_perfiles_italco_table(pdf_sources, excel_sources) -> pd.DataFrame:
    """Cruza los perfiles de la planilla ITALCO contra el histograma, por fecha.

    Equivalente ITALCO de :func:`build_perfiles_table`: el nivel se toma del Cargo
    (PDF) y de la CATEGORÍA (histograma), del histograma se excluyen los ítems de
    equipos (5.5) y servicios (5.6), y los ítems de 24 horas (5.1.1.4) se
    multiplican por 3 y se redondean antes de sumar.
    """
    if not isinstance(pdf_sources, (list, tuple)):
        pdf_sources = [pdf_sources]
    if not isinstance(excel_sources, (list, tuple)):
        excel_sources = [excel_sources]

    empty = pd.DataFrame(columns=["Fecha", "Nivel/Perfil", "PDF", "Excel", "Estado"])
    with tempfile.TemporaryDirectory(prefix="web_ui_perfiles_italco_") as tmp_dir:
        excel_paths = []
        for idx, excel_source in enumerate(excel_sources):
            if hasattr(excel_source, "getbuffer"):
                excel_path = os.path.join(tmp_dir, f"upload_{idx}.xlsx")
                with open(excel_path, "wb") as excel_handle:
                    excel_handle.write(excel_source.getbuffer())
            else:
                excel_path = str(excel_source)
            excel_paths.append(excel_path)

        partes_pdf = []
        for idx, pdf_source in enumerate(pdf_sources):
            if hasattr(pdf_source, "getbuffer"):
                pdf_path = os.path.join(tmp_dir, f"upload_{idx}.pdf")
                with open(pdf_path, "wb") as pdf_handle:
                    pdf_handle.write(pdf_source.getbuffer())
            else:
                pdf_path = str(pdf_source)
            df_parte = _extract_perfiles_italco_by_date(pdf_path)
            if not df_parte.empty:
                partes_pdf.append(df_parte)

        partes_excel = [_extract_excel_perfiles_italco_by_date(p) for p in excel_paths]
        partes_excel = [df for df in partes_excel if not df.empty]

        if not partes_pdf or not partes_excel:
            return empty

        df_pdf = pd.concat(partes_pdf, ignore_index=True).groupby(
            ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
        )["PDF"].sum()
        df_excel = pd.concat(partes_excel, ignore_index=True).groupby(
            ["FECHA", "PERFIL_NORM", "Nivel/Perfil"], as_index=False
        )["Excel"].sum()

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
        df_merge = df_merge[["Fecha", "Nivel/Perfil", "PDF", "Excel", "Estado"]]
        return df_merge.sort_values(["Fecha", "Nivel/Perfil"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Pagos (equipos / servicios / perfiles)
# ---------------------------------------------------------------------------

def process_pagos(pdf_files, excel_files, tipo, formato="tabarca"):
    """Compare PDF counts against the Excel history for equipos/servicios/perfiles.

    ``pdf_files`` / ``excel_files`` pueden ser un único archivo o una lista de
    varios; los conteos de todos los PDFs se acumulan, y la planilla histórica se
    arma sumando todos los Excel, antes de cruzar.

    ``formato`` (``tabarca`` / ``italco``) solo aplica a ``perfiles``: elige entre
    la planilla/cuadro TABARCA y la planilla de personal + histograma ITALCO.

    Returns ``(df_display, message)``: a results DataFrame (with a ``Fecha`` and
    ``Estado`` column, ready for the date filter + coloured table) or, when there
    is nothing to show, ``message`` explains why and ``df_display`` is ``None``.
    """
    if not isinstance(pdf_files, (list, tuple)):
        pdf_files = [pdf_files]
    if not isinstance(excel_files, (list, tuple)):
        excel_files = [excel_files]

    _debug_print(
        f"Inicio process_pagos. tipo={tipo}, formato={formato}, "
        f"pdfs={[getattr(f, 'name', 'N/A') for f in pdf_files]}, "
        f"excels={[getattr(f, 'name', 'N/A') for f in excel_files]}"
    )

    with tempfile.TemporaryDirectory(prefix="web_ui_") as tmp_dir:
        excel_paths = []
        for idx, excel_file in enumerate(excel_files):
            excel_path = os.path.join(tmp_dir, f"upload_{idx}.xlsx")
            with open(excel_path, "wb") as excel_handle:
                excel_handle.write(excel_file.getbuffer())
            excel_paths.append(excel_path)

        pdf_paths = []
        for idx, pdf_file in enumerate(pdf_files):
            pdf_path = os.path.join(tmp_dir, f"upload_{idx}.pdf")
            with open(pdf_path, "wb") as pdf_handle:
                pdf_handle.write(pdf_file.getbuffer())
            pdf_paths.append(pdf_path)

        rows = []

        if tipo == "perfiles":
            if str(formato).lower() == "italco":
                df_display = build_perfiles_italco_table(pdf_paths, excel_paths)
            else:
                df_display = build_perfiles_table(pdf_paths, excel_paths)
            _debug_print(f"Tabla perfiles ({formato}) construida. filas={len(df_display)}")
            if df_display.empty:
                return None, "No se encontraron perfiles en el PDF o no fue posible cruzarlos por fecha."
            return df_display, None
        else:
            partes_pdf = []
            for pdf_path in pdf_paths:
                df_parte = extraer_conteo_pdf(pdf_path, tipo)
                if df_parte is not None and not df_parte.empty:
                    partes_pdf.append(df_parte)
            if not partes_pdf:
                return None, f"No se encontraron elementos válidos de tipo {tipo} en el PDF."

            # Lado PDF: conteos por (fecha, clave robusta), guardando el nombre legible.
            df_pdf = pd.concat(partes_pdf, ignore_index=True)
            df_pdf["FECHA"] = pd.to_datetime(df_pdf["FECHA"], errors="coerce").dt.normalize()
            df_pdf["CLAVE"] = df_pdf["TIPO DE EQUIPO"].astype(str).map(clave_equipo)
            pdf_agg = df_pdf.groupby(["FECHA", "CLAVE"], as_index=False).agg(
                CANTIDAD=("CANTIDAD", "sum"),
                Nombre=("TIPO DE EQUIPO", "first"),
            )

            # Secciones del histograma a las que corresponden los PDF (5.5 equipos,
            # 5.6 servicios...), detectadas por el título de cada página. Así la
            # validación bidireccional solo abarca lo que el PDF debía reportar.
            prefijos = []
            for excel_path in excel_paths:
                prefijos = prefijos_seccion_pdf(excel_path, pdf_paths)
                if prefijos:
                    break
            _debug_print(f"Secciones del histograma detectadas para los PDF: {prefijos or '(sin filtro)'}")

            # Lado Excel: histograma en largo, filtrado a esas secciones (conserva
            # ceros) y sumado entre varios Excel.
            partes_excel = []
            for excel_path in excel_paths:
                try:
                    partes_excel.append(
                        leer_histograma_largo(excel_path, prefijos or None)
                    )
                except Exception:
                    print(f"[WARN][web_ui] No se pudo leer el histograma {excel_path}")
                    traceback.print_exc()
            if partes_excel:
                df_excel = pd.concat(partes_excel, ignore_index=True).groupby(
                    ["FECHA", "CLAVE"], as_index=False
                ).agg({"VALOR": "sum", "DESCRIPCION TARIFA": "first"})
            else:
                df_excel = pd.DataFrame(columns=["FECHA", "CLAVE", "VALOR", "DESCRIPCION TARIFA"])

            # Cruce bidireccional: todo lo del PDF debe estar en el Excel y viceversa.
            merged = pdf_agg.merge(df_excel, on=["FECHA", "CLAVE"], how="outer")
            for _, r in merged.iterrows():
                pdf_cnt = format_count(r["CANTIDAD"]) if pd.notna(r.get("CANTIDAD")) else 0
                excel_cnt = format_count(r["VALOR"]) if pd.notna(r.get("VALOR")) else 0
                # Ambos ausentes/cero: válido (p. ej. tarifa en Excel con valor 0 y
                # sin registro en el PDF). No aporta información -> no se muestra.
                if float(pdf_cnt or 0) == 0 and float(excel_cnt or 0) == 0:
                    continue
                nombre = r.get("Nombre")
                if not (isinstance(nombre, str) and nombre.strip()):
                    nombre = r.get("DESCRIPCION TARIFA")
                fecha = r["FECHA"]
                estado = "OK" if pdf_cnt == excel_cnt else "Valores diferentes"
                rows.append(
                    {
                        "Fecha": fecha.strftime("%Y-%m-%d") if hasattr(fecha, "strftime") else str(fecha),
                        "Servicio": nombre,
                        "PDF": pdf_cnt,
                        "Excel": excel_cnt,
                        "Estado": estado,
                    }
                )

            df_display = pd.DataFrame(rows)
            if not df_display.empty:
                df_display = df_display.sort_values(["Fecha", "Servicio"]).reset_index(drop=True)

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
        df_despr = procesar_desprendibles(dir_despr, transfer_format)
        df_trans = (
            procesar_transferencias(dir_trans, transfer_format)
            if os.listdir(dir_trans)
            else None
        )
        df_seg = (
            procesar_seguridad_social(dir_seg, transfer_format)
            if os.listdir(dir_seg)
            else None
        )

        reconcile_result = conciliar(df_despr, df_trans, df_seg)
        if isinstance(reconcile_result, tuple) and len(reconcile_result) == 2:
            df_transfers, df_seguridad = reconcile_result
        else:
            df_transfers = pd.DataFrame()
            df_seguridad = pd.DataFrame()

        if (df_transfers is None or df_transfers.empty) and (df_seguridad is None or df_seguridad.empty):
            return [], "No se encontraron registros o diferencias.", df_transfers, df_seguridad, []

        df_display_t = df_transfers.copy() if df_transfers is not None else pd.DataFrame()
        df_display_s = df_seguridad.copy() if df_seguridad is not None else pd.DataFrame()

        df_display_t = format_dataframe(df_display_t, formatear_valores)
        df_display_s = format_dataframe(df_display_s, formatear_valores)

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

def process_mano_obra(informe_files, ods_files, formato="tabarca"):
    """Cross the Informe de Costo against the ODS registry, per worker.

    ``informe_files`` / ``ods_files`` pueden ser un único archivo o una lista de
    varios Excel por lado; cada uno se lee por separado y se concatena antes de
    cruzar.

    ``formato`` (``tabarca`` / ``italco``) elige el layout del Informe: TABARCA
    usa la hoja ``Informe`` y ITALCO la progresión (filas de DIFERENCIA, nombre
    completo y fechas de actividades).

    Returns ``(table_html, message, df_result)`` where ``df_result`` keeps the
    list-valued comparison cells (one element = match, two = mismatch) so the
    Excel export can colour the same cells the HTML table highlights.
    """
    if not isinstance(informe_files, (list, tuple)):
        informe_files = [informe_files]
    if not isinstance(ods_files, (list, tuple)):
        ods_files = [ods_files]

    with tempfile.TemporaryDirectory(prefix="web_ui_mano_obra_") as tmp_dir:
        informe_paths = []
        for idx, informe_file in enumerate(informe_files):
            informe_path = os.path.join(tmp_dir, f"informe_{idx}.xlsx")
            with open(informe_path, "wb") as handle:
                handle.write(informe_file.getbuffer())
            informe_paths.append(informe_path)

        ods_paths = []
        for idx, ods_file in enumerate(ods_files):
            ods_path = os.path.join(tmp_dir, f"ods_{idx}.xlsx")
            with open(ods_path, "wb") as handle:
                handle.write(ods_file.getbuffer())
            ods_paths.append(ods_path)

        df_result = mano_obra.comparar_mano_obra(informe_paths, ods_paths, formato=formato)

    if df_result is None or df_result.empty:
        _debug_print("Mano de obra: ninguna persona cruzada entre Informe y ODS.")
        return None, "Ninguna persona del Informe se encontró en la ODS (cruce por documento).", df_result

    table_html = build_mano_obra_table(df_result)
    _debug_print(f"Tabla mano de obra construida. filas={len(df_result)}")
    return table_html, None, df_result
