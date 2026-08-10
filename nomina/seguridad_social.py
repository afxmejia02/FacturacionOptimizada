"""Lectura de las planillas de seguridad social (IBC) desde PDF."""
from __future__ import annotations

import os
import pandas as pd
import pdfplumber
import re

from .formato import _limpiar_numero
from .paginas import iter_paginas


def procesar_seguridad_social(folder_path, formato="tabarca"):
    """Procesa PDFs de seguridad social (IBC) según el formato indicado.

    Args:
        folder_path (str): ruta carpeta PDFs
        formato (str): "tabarca" o "italco"

    Returns:
        pd.DataFrame con columnas [archivo, cc, ibc]
    """
    formato = (formato or "tabarca").strip().lower()
    if formato == "italco":
        return _seguridad_social_italco(folder_path)
    return _seguridad_social_tabarca(folder_path)


def _seguridad_social_tabarca(folder_path):
    """
    Procesa PDFs de seguridad social y extrae:
    - CC
    - IBC únicos (solo el primer valor de la columna IBC por página)

    Args:
        folder_path (str): ruta carpeta PDFs

    Returns:
        pd.DataFrame
    """

    registros = []

    def _extraer_ibc_desde_tabla(tabla):
        """Obtiene todos los valores distintos de la columna IBC dentro de una tabla extraída."""
        if not tabla:
            return []

        columna_ibc = None
        fila_inicio = None

        for idx_fila, fila in enumerate(tabla):
            if not fila:
                continue
            for idx_col, celda in enumerate(fila):
                if celda and re.search(r"\bIBC\b", str(celda), re.IGNORECASE):
                    columna_ibc = idx_col
                    fila_inicio = idx_fila
                    break
            if columna_ibc is not None:
                break

        if columna_ibc is None:
            return []

        ibc_encontrados = []
        vistos = set()

        for fila in tabla[fila_inicio + 1:]:
            if not fila or columna_ibc >= len(fila):
                continue

            celda = fila[columna_ibc]
            if not celda:
                continue

            match_ibc = re.search(r"\$?\s*[\d\.,]+", str(celda))
            if not match_ibc:
                continue

            ibc = re.sub(r"[^\d]", "", match_ibc.group(0))
            if ibc and ibc not in vistos:
                vistos.add(ibc)
                ibc_encontrados.append(ibc)

        return ibc_encontrados

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, filename)

        with pdfplumber.open(path) as pdf:
            for page in iter_paginas(pdf):

                texto = page.extract_text()
                if not texto:
                    continue

                # =========================
                # 1. EXTRAER CC (más flexible: permite puntos/comas)
                # =========================
                match_cc = re.search(r'CC\s*[:\.-]?\s*([\d\.,]+)', texto, re.IGNORECASE)
                if not match_cc:
                    # fallback: any 6+ digit number
                    match_cc = re.search(r'\b(\d{6,})\b', texto)
                if not match_cc:
                    continue

                cc_raw = match_cc.group(1)
                cc = re.sub(r"[^\d]", "", cc_raw)

                # =========================
                # 2. EXTRAER IBC DESDE LA POSICIÓN DE LA COLUMNA EN LA TABLA
                # =========================
                ibc_set = set()

                tablas = []
                try:
                    tablas = page.extract_tables() or []
                except Exception:
                    tablas = []

                for tabla in tablas:
                    for ibc in _extraer_ibc_desde_tabla(tabla):
                        ibc_set.add(ibc)

                # Fallback: si no se pudo leer la tabla, usar el texto plano
                if not ibc_set:
                    lines = texto.splitlines()
                    for idx, line in enumerate(lines):
                        if re.search(r'\bIBC\b', line, re.IGNORECASE):
                            for follow in lines[idx: idx + 12]:
                                if not follow or follow.strip() == "":
                                    continue
                                match_ibc = re.search(r'\$?\s*[\d\.,]+', follow)
                                if match_ibc:
                                    valor = match_ibc.group(0)
                                    ibc = re.sub(r"[^\d]", "", valor)
                                    if ibc:
                                        ibc_set.add(ibc)
                            

                # =========================
                # 4. GUARDAR RESULTADOS
                # =========================
                for ibc in ibc_set:
                    registros.append({
                        "archivo": filename,
                        "cc": cc,
                        "ibc": ibc
                    })

                

    df = pd.DataFrame(registros)

    if not df.empty:
        df = df.drop_duplicates(subset=["cc", "ibc"])

    return df


def _seguridad_social_italco(folder_path):
    """
    Procesa la "Planilla Resumen" (aportes en línea) en formato ITALCO y
    extrae el documento y el IBC de pensión por fila.

    La columna del IBC depende del layout de la página: en la primera página
    las filas de personas empiezan en el índice 13 y el IBC está en la
    columna 26; en las páginas siguientes las filas válidas tienen 43 celdas
    y el IBC está en la columna 27.

    Returns:
        pd.DataFrame con columnas [archivo, cc, ibc]
    """
    registros = []

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, filename)
        with pdfplumber.open(path) as pdf:
            for page in iter_paginas(pdf):
                tabla = page.extract_table()
                if not tabla:
                    continue

                if page.page_number == 1:
                    filas = tabla[13:]
                    ibc_idx = 26
                else:
                    filas = [fila for fila in tabla if len(fila) == 43]
                    ibc_idx = 27

                for fila in filas:
                    if len(fila) <= ibc_idx:
                        continue

                    doc_raw = fila[2]
                    ibc_raw = fila[ibc_idx]
                    if not doc_raw or not ibc_raw:
                        continue

                    cc = re.sub(r"[^\d]", "", str(doc_raw))
                    ibc = _limpiar_numero(ibc_raw)
                    if not cc or ibc is None:
                        continue

                    registros.append({
                        "archivo": filename,
                        "cc": cc,
                        "ibc": int(ibc),
                    })

    df = pd.DataFrame(registros)

    if not df.empty:
        df = df.drop_duplicates(subset=["cc", "ibc"])

    return df
