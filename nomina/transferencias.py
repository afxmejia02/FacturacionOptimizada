"""Lectura de las transferencias bancarias (PDF) y emparejado de sus lineas."""
from __future__ import annotations

import os
import pandas as pd
import pdfplumber
import re

from .formato import _limpiar_numero, _normalizar_linea_ocr, _parsear_linea
from .depuracion import log as _log



# --- Detección robusta de renglones de soporte bancario (transferencias) -----
# La etiqueta de destino real ("PAGO NOMINA BCA") varía entre soportes y puede
# llegar con ruido de OCR. Solo se usa como señal de que la línea es un pago de
# nómina; el cruce no depende de su forma exacta. Tolera confusiones típicas de
# OCR: O/0, I/1, A/4.
_RE_NOMINA = re.compile(r"N[O0]M[I1]N[A4]", re.IGNORECASE)

# Importe monetario: admite miles con punto o coma y decimales opcionales.
#   1.234.567,89 | 1,234,567.89 | 5156483 | 913,143.00
_RE_IMPORTE = re.compile(r"\d{1,3}(?:[.,]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?")

# Documento del beneficiario (cédula/NIT): grupo de 5 a 15 dígitos.
_RE_DOCUMENTO = re.compile(r"\b(\d{5,15})\b")

# Cuenta / número de producto (opcional): grupo largo de 9+ dígitos.
_RE_PRODUCTO = re.compile(r"\b(\d{9,})\b")

# Número de factura: token (YYMMDD o YYYYMMDD) inmediatamente antes de "PAGO".
# Es la referencia de la quincena del soporte (p. ej. 250415 -> 2025-04-15), no
# la fecha real de consignación. Se usa para filtrar transferencias por periodo.
_RE_FACTURA = re.compile(r"(\d{6,8})\s+PAG", re.IGNORECASE)

def procesar_transferencias(folder_path, formato="tabarca"):
    """
    Extract transfer data from PDF files.

    Args:
        folder_path (str): Path to folder containing transfer PDFs
        formato (str): Transfer layout to parse ("tabarca" or "italco")

    Returns:
        pd.DataFrame: DataFrame with transfer information
    """
    formato = (formato or "tabarca").strip().lower()
    if formato == "italco":
        return _transferencias_italco(folder_path)
    return _transferencias_tabarca(folder_path)


def _transferencias_tabarca(folder_path):
    registros = []

    # Iterate PDF files in the folder and parse transfer lines
    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue
        path = os.path.join(folder_path, filename)
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if not texto:
                    continue

                for linea in texto.split("\n"):
                    linea = linea.strip()
                    # Detect lines with account numbers (10+ digits)
                    if re.match(r"\d{10,}", linea):
                        data = _parsear_linea(linea)
                        if data:
                            registros.append(data)

    df = pd.DataFrame(registros)

    if not df.empty and "Valor" in df.columns:
        # Normalize and convert Valor to integer using the existing cleaner
        def _to_int(v):
            try:
                if v is None:
                    return None
                num = _limpiar_numero(str(v))
                return int(num) if num is not None else None
            except Exception:
                return None

        df["Valor"] = df["Valor"].apply(_to_int)
        df["Valor"] = df["Valor"].fillna(0).astype("int64")

    return df


def _transferencias_italco(folder_path):
    registros = []
    total_lineas_pago = 0

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, filename)
        filas_archivo = 0

        with pdfplumber.open(path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                texto = page.extract_text() or ""
                if not texto:
                    _log(f"[transfer][{filename}] página {page_num} sin texto extraíble (¿escaneo sin OCR?).")
                    continue

                # Reconstruir líneas por palabras/coordenadas: separa columnas
                # contiguas que extract_text() pega (p. ej. documento+producto).
                page_rows = []
                for linea in _lineas_desde_palabras(page):
                    data = match_linea_transferencia(linea)
                    if data:
                        page_rows.append(data)

                if page_rows:
                    registros.extend(page_rows)
                    filas_archivo += len(page_rows)
                    total_lineas_pago += len(page_rows)
                    continue

                # Fallback: soporte tipo desprendible (una persona por página),
                # con la cédula tras "CC:" y el neto tras "Total Neto:".
                texto_plano = re.sub(r"\s+", " ", texto)
                cc_match = re.search(r"CC:\s*([\d.]+)", texto_plano, re.IGNORECASE)
                neto_match = re.search(r"Total Neto:\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                if cc_match and neto_match:
                    cc = re.sub(r"[^\d]", "", cc_match.group(1))
                    neto = _limpiar_numero(neto_match.group(1))
                    registros.append(
                        {
                            "Cuenta": None,
                            "Tipo": "ITALCO",
                            "Documento": cc,
                            "Nombre": None,
                            "Valor": neto,
                            "Fecha": None,
                        }
                    )
                    filas_archivo += 1
                    total_lineas_pago += 1
                else:
                    _log(
                        f"[transfer][{filename}] página {page_num}: sin renglones de "
                        f"nómina ni patrón de respaldo (CC/Total Neto)."
                    )

        _log(f"[transfer][{filename}] {filas_archivo} valor(es) de transferencia extraído(s).")

    df = pd.DataFrame(registros)
    _log(f"[transfer] Total transferencias extraídas: {total_lineas_pago} en {len(df)} registros.")

    if not df.empty and "Valor" in df.columns:
        def _to_int(v):
            try:
                if v is None:
                    return None
                num = _limpiar_numero(str(v))
                return int(num) if num is not None else None
            except Exception:
                return None

        df["Valor"] = df["Valor"].apply(_to_int)
        df["Valor"] = df["Valor"].fillna(0).astype("int64")

    if not df.empty:
        dedupe_cols = [col for col in ("Documento", "Cuenta", "Valor", "Fecha") if col in df.columns]
        if dedupe_cols:
            df = df.drop_duplicates(subset=dedupe_cols)

    return df


def match_linea_transferencia(linea):
    """Extrae ``{Documento, Cuenta, Nombre, Valor}`` de una línea de soporte.

    Diseño robusto y genérico (no atado a un layout concreto). Tolera:

    - **ausencia/presencia de la columna de fecha** (el bug original: el
      patrón exigía una fecha de 8 dígitos que muchos soportes no traen);
    - cualquier **formato monetario** (miles con punto o coma, con/sin
      decimales) vía :func:`_limpiar_numero`;
    - **espacios y OCR ruidoso** (se normaliza la línea antes de leerla);
    - **variantes de la etiqueta de destino** ("PAGO NOMINA BCA",
      "PAGO DE NOMINA", "PAGONOMINA"...): basta con que aparezca "NOMINA".

    Estrategia por proximidad de campos en el renglón:
      ``[productos] <nombre> <documento> [cuenta] [fechaPago] <factura> <...NOMINA...> <valor> [ods]``
      - documento = primer grupo de 5-15 dígitos (cédula/NIT del beneficiario);
      - valor = primer importe DESPUÉS de la etiqueta de destino (evita tomar
        la columna "ods"/consecutivo que algunos soportes ponen al final);
      - cuenta = primer grupo largo (9+ dígitos) tras el documento (opcional);
      - fecha = el número de factura (YYMMDD) justo antes de "PAGO", que indica
        la quincena del soporte (se usa luego para filtrar por periodo).

    Hay un segundo layout (la "consulta de pagos a terceros" del banco) cuyos
    renglones **no traen la etiqueta NÓMINA** ni la fecha-factura de quincena
    (solo la fecha de consignación, que puede ser de otro mes). Esos renglones
    se devuelven como **candidatos** (``EsNomina=False``): se extraen documento
    y valor, pero solo se aceptan en la conciliación si su valor coincide con
    un neto del desprendible (los que no, pueden ser de otra quincena).

    Devuelve ``None`` si la línea no parece un renglón de pago.
    """
    plano = _normalizar_linea_ocr(linea)
    if not plano:
        return None

    # 1) Documento del beneficiario: primer grupo de 5-15 dígitos.
    m_doc = _RE_DOCUMENTO.search(plano)
    if not m_doc:
        return None
    documento = m_doc.group(1)
    nombre = plano[: m_doc.start()].strip(" -")
    # Quitar un eventual número de "Productos" al inicio del renglón.
    nombre = re.sub(r"^[\d\s]+", "", nombre).strip(" -") or None

    # 2) ¿Trae la etiqueta NÓMINA? (señal tolerante a OCR). Si la trae, es una
    #    transferencia confiable; si no, es un candidato a validar por valor.
    m_nomina = _RE_NOMINA.search(plano)
    es_nomina = bool(m_nomina)

    # 3) Valor.
    if m_nomina:
        # Confiable: primer importe DESPUÉS de la etiqueta de destino. Así no se
        # confunde con la columna "ods" (1-3 dígitos) al final del renglón.
        importes_post = _RE_IMPORTE.findall(plano[m_nomina.end():])
        if importes_post:
            valor = _limpiar_numero(importes_post[0])
        else:
            importes = _RE_IMPORTE.findall(plano)
            valor = _limpiar_numero(importes[-1]) if importes else None
    else:
        # Candidato: último importe monetario (con separador de miles/decimales)
        # tras el documento, que es la columna "Vr. Pago" en este layout. Exigir
        # el separador evita tomar identificadores/consecutivos como valor.
        importes = [t for t in _RE_IMPORTE.findall(plano[m_doc.end():]) if re.search(r"[.,]", t)]
        valor = _limpiar_numero(importes[-1]) if importes else None
    if valor is None:
        return None

    # 4) Cuenta/producto (opcional): primer grupo largo tras el documento.
    cola = plano[m_doc.end():]
    m_cta = _RE_PRODUCTO.search(cola)
    cuenta = m_cta.group(1) if m_cta else None

    # 5) Fecha de la factura (referencia de la quincena): token antes de "PAGO".
    fecha = None
    m_fact = _RE_FACTURA.search(plano)
    if m_fact:
        factura = m_fact.group(1)
        formato = "%Y%m%d" if len(factura) == 8 else "%y%m%d"
        parsed = pd.to_datetime(factura, format=formato, errors="coerce")
        fecha = None if pd.isna(parsed) else parsed

    return {
        "Cuenta": cuenta,
        "Tipo": "ITALCO",
        "Documento": documento,
        "Nombre": nombre,
        "Valor": valor,
        "Fecha": fecha,
        "EsNomina": es_nomina,
    }


def _lineas_desde_palabras(page, x_tolerance=1.5, y_tolerance=3.0):
    """Reconstruye las líneas de la página a partir de las palabras y sus
    coordenadas (x, y), no de ``extract_text()``.

    Es necesario porque en algunos soportes columnas contiguas (p. ej. la
    columna "Productos" y el documento, o el documento y el número de
    producto) quedan pegadas o entrelazadas en ``extract_text()``, lo que
    rompe la lectura del documento. Agrupando por línea (``top``) y ordenando
    por ``x0`` se recuperan los campos como tokens separados.

    Si no hay palabras (página sin capa de texto), cae a ``extract_text()``.
    """
    try:
        words = page.extract_words(x_tolerance=x_tolerance, use_text_flow=False)
    except Exception:
        words = []
    if not words:
        return (page.extract_text() or "").split("\n")

    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lineas = []
    actual = []
    base_top = None
    for w in words:
        if base_top is not None and abs(w["top"] - base_top) > y_tolerance:
            lineas.append(" ".join(c["text"] for c in sorted(actual, key=lambda c: c["x0"])))
            actual = []
            base_top = None
        actual.append(w)
        if base_top is None:
            base_top = w["top"]
    if actual:
        lineas.append(" ".join(c["text"] for c in sorted(actual, key=lambda c: c["x0"])))
    return lineas
