"""Lectura de los desprendibles de nomina (PDF) en formato TABARCA e ITALCO."""
from __future__ import annotations

import os
import pandas as pd
import pdfplumber
import re

from .formato import _limpiar_numero
from .paginas import iter_paginas


def procesar_desprendibles(folder_path, formato="tabarca"):
    """
    Extrae datos de desprendibles desde archivos PDF.

    Args:
        folder_path (str): Ruta de la carpeta con PDFs de desprendibles
        formato (str): Formato del desprendible a interpretar ("tabarca" o "italco")

    Returns:
        pd.DataFrame: DataFrame con las columnas [Identificacion, Neto, Devengado, Cuenta]
    """
    formato = (formato or "tabarca").strip().lower()
    if formato == "italco":
        return _desprendibles_italco(folder_path)
    return _desprendibles_tabarca(folder_path)


def _desprendibles_tabarca(folder_path):
    registros = []
    
    # Buscar PDFs de desprendibles (normalmente nombrados por mes/año)
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            path = os.path.join(folder_path, filename)
            
            with pdfplumber.open(path) as pdf:
                for page in iter_paginas(pdf):
                    texto = page.extract_text()
                    if not texto:
                        continue
                    
                    # Separar por el marcador del bloque del desprendible
                    bloques = texto.split("Comprobante de Nómina")
                    
                    for bloque in bloques:
                        # Extraer identificación (formato con puntos: 123.456.789)
                        id_match = re.search(r"\b\d{1,3}(?:[.,]\d{3}){1,3}\b", bloque)
                        cuenta_match = re.search(r"Cuenta No\s*(\d{6,})\b", bloque)
                        
                        # Extraer el valor neto a pagar
                        # Extraer el valor neto a pagar (más flexible, case-insensitive)
                        neto_match = re.search(
                            r"Neto(?:\s+a\s+pagar)?[:\s].*?\$\s*([\d\.,]+)",
                            bloque,
                            re.IGNORECASE | re.DOTALL
                        )

                        # Extraer Devengado: buscar 'Devengado' primero, si no, intentar 'TOTALES'
                        #si en la columna SALDOS hay

                        deven_match = re.search(
                                r"TOTALES[:\s].*?\$\s*([\d\.,]+)",
                                bloque,
                                re.IGNORECASE | re.DOTALL
                            )
                        
                        aux_match = re.search(r"AUXILIO DE LOCALIZACION\s+([\d.,]+)\s+([\d.,]+)", bloque)
                        pri_match = re.search(r"PRIMA LEGAL DE SERVICIOS JUNIO\s+([\d.,]+)\s+([\d.,]+)", bloque)
                        vac_match = re.search(r"VACACIONES INDEMNIZADAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                        ces_match = re.search(r"CESANTIAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                        int_ces_match= re.search(r"INTERESES DE CESANTIAS\s+([\d.,]+)\s+([\d.,]+)", bloque)
                        

                        
                        if id_match and neto_match:
                            identificacion = id_match.group()
                            identificacion = identificacion.replace(".", "").replace(",", "")
                            neto = _limpiar_numero(neto_match.group(1))
                            devengado = _limpiar_numero(deven_match.group(1)) if deven_match else None
                            dtos= []
                            for m in [aux_match, pri_match, vac_match, ces_match, int_ces_match]:
                                if m:
                                    valor = _limpiar_numero(m.group(2))
                                    if valor is not None:
                                        dtos.append(valor)
                            if devengado is not None and sum(dtos) is not None:
                                devengado -= sum(dtos)
                            cuenta = cuenta_match.group(1) if cuenta_match else None
                            
                            registros.append({
                                "Identificacion": identificacion,
                                "Neto": neto,
                                "Devengado": devengado,
                                "Cuenta": cuenta
                            })
    
    df = pd.DataFrame(registros)
    
    if not df.empty:
        # Eliminar puntos de la identificación
        df["Identificacion"] = df["Identificacion"].str.replace(".", "", regex=False)
        # Convertir Neto a entero
        df["Neto"] = df["Neto"].fillna(0).astype("int64")
        # Convertir Devengado a entero
        if "Devengado" in df.columns:
            df["Devengado"] = df["Devengado"].fillna(0).astype("int64")
    return df


def _desprendibles_italco(folder_path):
    """
    Extrae datos de los comprobantes de pago de nómina en formato ITALCO.

    A diferencia del formato TABARCA, cada página es un comprobante con la
    cédula tras ``CC:`` y el neto tras ``Total Neto:`` (sin símbolo ``$``).

    Returns:
        pd.DataFrame: DataFrame con las columnas [Identificacion, Neto, Devengado, Cuenta]
    """
    registros = []

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue

        path = os.path.join(folder_path, filename)
        with pdfplumber.open(path) as pdf:
            for page in iter_paginas(pdf):
                texto = page.extract_text() or ""
                if not texto:
                    continue

                texto_plano = re.sub(r"\s+", " ", texto)
                # Documento: preferir "CC: <num>"; si el CC viene vacío (algunas
                # plantillas lo dejan en blanco), usar "Documento <num>" de la
                # línea del trabajador.
                cc_match = re.search(r"CC:\s*(\d[\d.,]*)", texto_plano, re.IGNORECASE)
                doc_match = re.search(r"\bDocumento\b\s*:?\s*(\d[\d.,]*)", texto_plano, re.IGNORECASE)
                fuente_doc = cc_match or doc_match
                # Los valores pueden traer símbolo "$" (p. ej. "Total Neto: $3,675,627").
                neto_match = re.search(r"Total Neto:\s*\$?\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                cuenta_match = re.search(r"CUENTA:\s*(\d+)", texto_plano, re.IGNORECASE)
                # El devengado en ITALCO es el TOTAL INGRESOS (no el Total Neto).
                dev_match = re.search(r"TOTAL INGRESOS\s*\$?\s*([\d,\.]+)", texto_plano, re.IGNORECASE)
                # Periodo de la quincena: "Periodo: 2025-04-01 al 2025-04-15".
                # Se usa para descartar transferencias de otras quincenas/meses.
                periodo_match = re.search(
                    r"Periodo:?\s*(\d{4}-\d{1,2}-\d{1,2})\s*al\s*(\d{4}-\d{1,2}-\d{1,2})",
                    texto_plano,
                    re.IGNORECASE,
                )

                if not (fuente_doc and neto_match):
                    continue

                identificacion = re.sub(r"[^\d]", "", fuente_doc.group(1))
                neto = _limpiar_numero(neto_match.group(1))
                devengado = _limpiar_numero(dev_match.group(1)) if dev_match else None
                cuenta = cuenta_match.group(1) if cuenta_match else None
                periodo_inicio = (
                    pd.to_datetime(periodo_match.group(1), errors="coerce")
                    if periodo_match else pd.NaT
                )
                periodo_fin = (
                    pd.to_datetime(periodo_match.group(2), errors="coerce")
                    if periodo_match else pd.NaT
                )

                if not identificacion:
                    continue

                registros.append({
                    "Identificacion": identificacion,
                    "Neto": neto,
                    "Devengado": devengado,
                    "Cuenta": cuenta,
                    "PeriodoInicio": periodo_inicio,
                    "PeriodoFin": periodo_fin,
                })

    df = pd.DataFrame(registros)

    if not df.empty:
        df["Neto"] = df["Neto"].fillna(0).astype("int64")
        df["Devengado"] = df["Devengado"].fillna(0).astype("int64")
    return df
