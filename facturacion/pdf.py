"""Extraccion de conteos (perfiles, equipos y servicios) desde el informe PDF."""
from __future__ import annotations

import re
from collections import Counter

import pandas as pd
import pdfplumber

from .depuracion import debug as _debug
from .paginas import iter_paginas
from .normalizacion import (
    _buscar_indice_columna,
    _normalizar_texto_equipo,
    es_celda_vacia,
    limpiar_nombre_equipo,
    normalizar_busqueda,
    normalizar_fecha,
    normalizar_perfil,
    parsear_cantidad,
    parsear_observacion_perfil,
)


_ETIQUETAS_EQUIPO = ("equipo", "tipo de equipo", "tipo equipo")
_ETIQUETAS_SERVICIO = ("servicio", "tipo de servicio", "servicios")

def extraer_valor_etiqueta(tablas, etiquetas_objetivo):
    """Devuelve el valor asociado a una etiqueta tipo ``EQUIPO:`` en las tablas.

    La celda debe **ser** la etiqueta (igualdad, ignorando ``:`` final), no
    solo contenerla: de lo contrario ``"equipo"`` coincidiría con la palabra
    ``"EQUIPOS"`` dentro de un texto largo (p. ej. la descripción de la orden
    de servicio), y se tomaría como valor la celda equivocada.

    Soporta dos disposiciones:
      - etiqueta y valor en celdas separadas (``EQUIPO:`` | ``CAMIÓN-GRÚA…``);
      - etiqueta y valor en la misma celda (``EQUIPO: CAMIÓN-GRÚA…``).
    """
    # Etiquetas normalizadas (sin acentos, minúsculas, sin espacios ni ':').
    etiquetas_norm = tuple(
        normalizar_busqueda(etiqueta).replace(" ", "").rstrip(":")
        for etiqueta in etiquetas_objetivo
    )

    for tabla in tablas:
        for row in tabla:
            for i, cell in enumerate(row):
                cell_norm = normalizar_busqueda(cell).replace(" ", "")
                if not cell_norm:
                    continue

                # (a) La celda ES exactamente la etiqueta: el valor está en la
                # siguiente celda no vacía de la fila.
                if cell_norm.rstrip(":") in etiquetas_norm:
                    for next_cell in row[i + 1 :]:
                        if next_cell and str(next_cell).strip():
                            return limpiar_nombre_equipo(next_cell)

                # (b) Etiqueta y valor en la misma celda ("EQUIPO: <valor>").
                if any(cell_norm.startswith(etiqueta + ":") for etiqueta in etiquetas_norm):
                    partes = str(cell).split(":", 1)
                    if len(partes) == 2 and partes[1].strip():
                        return limpiar_nombre_equipo(partes[1])
    return None

def _extraer_fecha_reporte(page_text, tablas):
    patron_fecha = (
        r"(\d{1,2}\s+de\s+[a-záéíóúñ]+\s+de\s+\d{4}|"
        r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|"
        r"\d{4}-\d{1,2}-\d{1,2})"
    )

    fecha_match = re.search(rf"fecha\s*:?.*?{patron_fecha}", page_text, flags=re.IGNORECASE)
    if fecha_match:
        return normalizar_fecha(fecha_match.group(1))

    fecha_match = re.search(patron_fecha, page_text, flags=re.IGNORECASE)
    if fecha_match:
        return normalizar_fecha(fecha_match.group(1))

    for tabla in tablas:
        for row in tabla:
            for cell in row:
                if not cell:
                    continue
                cell_text = str(cell)
                if "fecha" not in normalizar_busqueda(cell_text):
                    continue
                fecha_match = re.search(patron_fecha, cell_text, flags=re.IGNORECASE)
                if fecha_match:
                    return normalizar_fecha(fecha_match.group(1))

    return None

def extraer_perfiles_pdf(path_planilla):
    conteo = Counter()
    fecha_reporte = None

    with pdfplumber.open(path_planilla) as pdf:
        for page in iter_paginas(pdf):
            for tabla in page.extract_tables() or []:
                if not tabla or len(tabla) <= 7:
                    continue

                header = tabla[6]
                
                header_norm = [normalizar_busqueda(celda).replace(" ", "") if celda else "" for celda in header]
                if "nivel/perfil" not in header_norm:
                    continue

                idx_perfil = header_norm.index("nivel/perfil")
                if header:
                    fecha_detectada = normalizar_fecha(header[-1])
                    if fecha_detectada is not None:
                        fecha_reporte = fecha_detectada
                tabla_info = str(tabla[4][2])
                for row in tabla[7:]:
                    if len(row) <= idx_perfil:
                        continue
                    perfil = row[idx_perfil]
                    observacion = row[-1]
                    
                    tabla_info_upper = normalizar_busqueda(tabla_info).upper()
                    if "GLOBAL" in tabla_info_upper or "NO FACTURABLE" in tabla_info_upper:
                        continue

                    celda_validacion = row[7] if len(row) > 7 else None
                    if es_celda_vacia(celda_validacion):
                        _debug(
                            f"Fila omitida por row[7] vacio: {repr(celda_validacion)} | perfil={repr(perfil)} | observacion={repr(observacion)}"
                        )
                        continue

                    # Interpretar Observaciones (recategorización, "E y F",
                    # "NO FACTURABLE" y "24h", que pueden coexistir y en cualquier orden).
                    recategorizado, es_ef, no_facturable, es_24h_obs = (
                        parsear_observacion_perfil(observacion)
                    )
                    if no_facturable:
                        continue

                    if recategorizado:
                        fuente = recategorizado
                    elif es_ef or es_24h_obs or es_celda_vacia(observacion):
                        # "E y F", "24 horas" o sin observación: el nivel es el
                        # de la columna (no la última palabra de la observación).
                        fuente = perfil.strip() if isinstance(perfil, str) else perfil
                    else:
                        # Otra observación no reconocida: comportamiento previo.
                        fuente = str(observacion).split()[-1]

                    # 24 horas (al inicio de la hoja o en observaciones): 1/3 por
                    # persona, salvo "E y F" (cuenta como 1).
                    es_24h = ("24" in tabla_info) or es_24h_obs
                    cantidad = 1 / 3 if (es_24h and not es_ef) else 1
                    if fuente:
                        conteo[normalizar_perfil(fuente)] += cantidad
                                         
    return conteo, fecha_reporte

def extraer_registros_etiqueta(tablas, etiquetas):
    """Registros ``[FECHA, TIPO DE EQUIPO, CANTIDAD]`` de una página cuyo
    'tipo' (equipo o servicio) está en una etiqueta tipo ``EQUIPO:`` /
    ``SERVICIO:`` y cuyo detalle trae columnas FECHA y CANTIDAD por fila.

    Equipos y servicios (formato vigente) comparten esta estructura; por eso
    un mismo extractor sirve para ambos y para PDFs que mezclan los dos.
    """
    tipo_valor = extraer_valor_etiqueta(tablas, etiquetas)
    if not tipo_valor:
        return []

    registros = []
    for tabla in tablas:
        if not tabla or len(tabla) < 3:
            continue

        header_idx = idx_fecha = idx_cantidad = None
        for i, row in enumerate(tabla[:12]):
            idx_fecha = _buscar_indice_columna(row, ("fecha", "dia"))
            idx_cantidad = _buscar_indice_columna(row, ("cant", "cantidad"))
            if idx_fecha is not None and idx_cantidad is not None:
                header_idx = i
                break
        if header_idx is None:
            continue

        for row in tabla[header_idx + 1 :]:
            if len(row) <= max(idx_fecha, idx_cantidad):
                continue
            fecha = normalizar_fecha(row[idx_fecha])
            cantidad = parsear_cantidad(row[idx_cantidad])
            if not fecha or cantidad is None:
                continue
            registros.append(
                {"FECHA": fecha, "TIPO DE EQUIPO": tipo_valor, "CANTIDAD": cantidad}
            )
    return registros

def _extraer_registros_servicios_legacy(page_text, tablas):
    """Formato antiguo de servicios: una fecha de reporte por página y el
    servicio en una columna del detalle (no en una etiqueta)."""
    fecha_reporte = _extraer_fecha_reporte(page_text, tablas)
    if fecha_reporte is None:
        return []

    registros = []
    for tabla in tablas:
        if not tabla or len(tabla) < 3:
            continue

        header_idx = idx_tipo = idx_cantidad = None
        for i, row in enumerate(tabla[:12]):
            idx_tipo = _buscar_indice_columna(row, ("tipo de equipo", "tipo equipo", "servicio"))
            idx_cantidad = _buscar_indice_columna(row, ("cant", "cantidad"))
            if idx_tipo is not None and idx_cantidad is not None:
                header_idx = i
                break
        if header_idx is None:
            continue

        for row in tabla[header_idx + 1 :]:
            if len(row) <= max(idx_tipo, idx_cantidad):
                continue
            tipo = _normalizar_texto_equipo(row[idx_tipo])
            cantidad = parsear_cantidad(row[idx_cantidad])
            if not isinstance(tipo, str) or not tipo.strip() or cantidad is None:
                continue
            registros.append(
                {"FECHA": fecha_reporte, "TIPO DE EQUIPO": tipo, "CANTIDAD": cantidad}
            )
    return registros

def _extraer_conteo_pdf_detallado(path_planilla, tipo_formato):
    """Extrae registros de un PDF de equipos y/o servicios.

    - ``equipos`` / ``servicios``: reconoce su etiqueta (``EQUIPO:`` /
      ``SERVICIO:``) con el detalle por fila.
    - ``equipos_servicios``: reconoce AMBAS, de modo que un solo PDF que
      mezcle páginas de equipos y de servicios se procesa de una vez.

    Para servicios y el modo combinado hay un *fallback* al formato antiguo
    (fecha de reporte + columna de servicio) cuando una página no trae la
    etiqueta.
    """
    # Normalizar alias: "equipos y servicios" / "todos" -> "equipos_servicios".
    clave = str(tipo_formato).lower().strip().replace(" y ", "_").replace(" ", "_")

    if clave == "equipos":
        etiquetas, usar_legacy = _ETIQUETAS_EQUIPO, False
    elif clave == "servicios":
        etiquetas, usar_legacy = _ETIQUETAS_SERVICIO, True
    elif clave in ("equipos_servicios", "todos"):
        etiquetas, usar_legacy = _ETIQUETAS_EQUIPO + _ETIQUETAS_SERVICIO, True
    else:
        raise ValueError(f"Unknown extraction type: {tipo_formato}")

    registros = []
    with pdfplumber.open(path_planilla) as pdf:
        for page in iter_paginas(pdf):
            tablas = page.extract_tables() or []
            regs = extraer_registros_etiqueta(tablas, etiquetas)
            if not regs and usar_legacy:
                regs = _extraer_registros_servicios_legacy(page.extract_text() or "", tablas)
            registros.extend(regs)

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    return df.groupby(["FECHA", "TIPO DE EQUIPO"], as_index=False)["CANTIDAD"].sum()

def _extraer_equipos_pdf(path_planilla):
    return _extraer_conteo_pdf_detallado(path_planilla, "equipos")

def _extraer_servicios_pdf(path_planilla):
    return _extraer_conteo_pdf_detallado(path_planilla, "servicios")

def _extraer_equipos_servicios_pdf(path_planilla):
    return _extraer_conteo_pdf_detallado(path_planilla, "equipos_servicios")

def extraer_conteo_pdf(path_planilla, tipo_extraccion="equipos"):
    clave = str(tipo_extraccion).lower().strip().replace(" y ", "_").replace(" ", "_")
    if clave == "perfiles":
        return extraer_perfiles_pdf(path_planilla)
    if clave == "equipos":
        return _extraer_equipos_pdf(path_planilla)
    if clave == "servicios":
        return _extraer_servicios_pdf(path_planilla)
    if clave in ("equipos_servicios", "todos"):
        return _extraer_equipos_servicios_pdf(path_planilla)
    raise ValueError(f"Unknown extraction type: {tipo_extraccion}")
