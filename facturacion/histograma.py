"""Lectura del Excel historico (histograma) de facturacion."""
from __future__ import annotations

import datetime as dt
import os

import pandas as pd
import pdfplumber

from .normalizacion import clave_equipo, normalizar_busqueda


def leer_excel_facturacion(path_hist):
    """Lee el Excel detectando la fila de encabezado real.

    Algunos cuadros traen filas de título antes de la cabecera
    ``COD. TAR. | DESCRIPCION TARIFA | ...``; se localiza esa fila y se usa
    como encabezado en vez de asumir la primera fila.
    """
    crudo = pd.read_excel(path_hist, header=None)
    fila_encabezado = 0
    for i in range(min(15, len(crudo))):
        celdas = {normalizar_busqueda(v) for v in crudo.iloc[i].tolist()}
        if "descripcion tarifa" in celdas:
            fila_encabezado = i
            break
    return pd.read_excel(path_hist, header=fila_encabezado)

def col_codigo_tarifa(df):
    """Devuelve el nombre de la columna de código de tarifa (``COD. TAR.``)."""
    return next(
        (c for c in df.columns if "cod" in normalizar_busqueda(str(c))),
        None,
    )

def prefijos_seccion_pdf(path_hist, paths_pdf):
    """Detecta a qué secciones del histograma corresponden los PDF.

    Cada página de los PDF trae como **título** (primera línea) la sección a
    la que pertenece (p. ej. "...ELEMENTOS, HERRAMIENTAS Y EQUIPOS
    TRANSVERSALES" o "...OBRAS O SERVICIOS TÍPICOS"). Ese título se casa con la
    descripción del encabezado de sección del histograma y se devuelve su
    ``COD. TAR.`` (p. ej. ``5.5`` para equipos, ``5.6`` para servicios). Así la
    validación bidireccional solo abarca lo que el PDF debía reportar y no
    otras secciones (perfiles, etc.).

    Devuelve la lista de prefijos de código (sin duplicar). Vacía si no logra
    emparejar ningún título (en ese caso el llamador no filtra por sección).
    """
    if isinstance(paths_pdf, (str, os.PathLike)):
        paths_pdf = [paths_pdf]

    titulos = set()
    for path in paths_pdf:
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    for linea in texto.splitlines()[:1]:  # título = 1ª línea
                        if linea.strip():
                            titulos.add(normalizar_busqueda(linea))
        except Exception:
            continue
    if not titulos:
        return []

    df_hist = leer_excel_facturacion(path_hist)
    col_cod = col_codigo_tarifa(df_hist)
    if col_cod is None or "DESCRIPCION TARIFA" not in df_hist.columns:
        return []

    prefijos = []
    for _, row in df_hist.iterrows():
        desc = row.get("DESCRIPCION TARIFA")
        cod = row.get(col_cod)
        if not isinstance(desc, str) or pd.isna(cod):
            continue
        desc_norm = normalizar_busqueda(desc)
        # Solo descripciones de sección (suficientemente largas) contenidas en
        # algún título de página del PDF.
        if len(desc_norm) >= 10 and any(desc_norm in t for t in titulos):
            prefijos.append(str(cod).strip())
    return list(dict.fromkeys(prefijos))

def leer_histograma_largo(path_hist, prefijos_cod=None):
    """Histograma en formato largo por fecha, para validación bidireccional.

    Devuelve un DataFrame con columnas ``FECHA``, ``DESCRIPCION TARIFA``,
    ``CLAVE`` y ``VALOR`` (un registro por tarifa y fecha). **Conserva los
    ceros** (un valor 0 en el Excel sin registro en el PDF es válido). Si
    ``prefijos_cod`` se indica, solo se conservan las tarifas cuyo
    ``COD. TAR.`` pertenece a esas secciones (p. ej. ``5.5`` / ``5.6``).

    El valor se toma **tal cual** del Excel (sin conversión de unidades).
    """
    df_hist = leer_excel_facturacion(path_hist)
    if "DESCRIPCION TARIFA" not in df_hist.columns:
        raise KeyError("Excel file missing 'DESCRIPCION TARIFA' column.")

    df_niveles = df_hist[df_hist["DESCRIPCION TARIFA"].notna()].copy()

    if prefijos_cod:
        col_cod = col_codigo_tarifa(df_niveles)
        if col_cod is not None:
            cods = df_niveles[col_cod].astype(str).str.strip()
            mask = pd.Series(False, index=df_niveles.index)
            for pref in prefijos_cod:
                pref = str(pref).strip()
                mask |= (cods == pref) | cods.str.startswith(pref + ".")
            df_niveles = df_niveles[mask]

    cols_fecha = [c for c in df_niveles.columns if isinstance(c, (pd.Timestamp, dt.datetime))]
    if not cols_fecha:
        raise ValueError("No date columns detected in Excel file.")

    cols_id = [c for c in df_niveles.columns if c not in cols_fecha]
    largo = df_niveles.melt(id_vars=cols_id, value_vars=cols_fecha, var_name="FECHA", value_name="VALOR")
    largo["FECHA"] = pd.to_datetime(largo["FECHA"], errors="coerce").dt.normalize()
    largo = largo[largo["VALOR"].notna()]
    largo["VALOR"] = pd.to_numeric(largo["VALOR"], errors="coerce")
    largo = largo[largo["VALOR"].notna()]
    largo["CLAVE"] = largo["DESCRIPCION TARIFA"].apply(clave_equipo)
    return largo.groupby(["FECHA", "CLAVE"], as_index=False).agg(
        {"VALOR": "sum", "DESCRIPCION TARIFA": "first"}
    )
