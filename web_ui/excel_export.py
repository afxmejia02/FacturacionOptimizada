"""Construye un Excel (.xlsx) con los resultados para guardar/descargar.

Espeja la lógica de ``pdf_export``: una hoja por sección, encabezado con estilo
y celdas inconsistentes resaltadas en rojo. Las celdas tipo lista (mano de obra)
se aplanan: un elemento -> el valor; dos -> ``Informe: a | Lista ODS: b`` (celda
resaltada). Mantiene todo el estilo aquí para que ``app.py`` solo escriba bytes.
"""
from __future__ import annotations

import io
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_DIFF_FILL = PatternFill("solid", fgColor="F8D7DA")
_DIFF_FONT = Font(color="721C24", bold=True)
_OK_FILL = PatternFill("solid", fgColor="D4EDDA")
_OK_FONT = Font(color="155724")
_WRAP_TOP = Alignment(vertical="top", wrap_text=True)


def _texto_celda(value, source_labels):
    """``(texto, es_diferencia)`` para una celda; aplana listas (mano de obra)."""
    if isinstance(value, (list, tuple)):
        items = ["" if v is None else str(v) for v in value]
        if len(items) <= 1:
            return (items[0] if items else ""), False
        etiquetas = source_labels or ("Informe", "Lista ODS")
        return f"{etiquetas[0]}: {items[0]} | {etiquetas[1]}: {items[1]}", True

    if value is None:
        return "", False
    try:
        if pd.isna(value):
            return "", False
    except (TypeError, ValueError):
        pass
    return str(value), False


def _nombre_hoja(subtitulo, usados):
    """Nombre de hoja válido (<=31 chars, sin caracteres prohibidos, único)."""
    base = re.sub(r"[\\/*?:\[\]]", " ", str(subtitulo or "Resultados")).strip() or "Resultados"
    base = base[:31]
    nombre = base
    i = 2
    while nombre in usados:
        sufijo = f" ({i})"
        nombre = base[: 31 - len(sufijo)] + sufijo
        i += 1
    usados.add(nombre)
    return nombre


def build_results_excel(titulo, archivos, secciones, source_labels=None):
    """Renderiza el Excel de resultados y devuelve sus bytes.

    Mismos parámetros que :func:`pdf_export.build_results_pdf`:
    - ``titulo``: encabezado.
    - ``archivos``: dict ``{"PDF": "...", ...}`` (puede traer listas de nombres).
    - ``secciones``: lista de ``(subtitulo, DataFrame)``; una hoja por sección.
    - ``source_labels``: par para las celdas con dos valores (mano de obra).
    """
    wb = Workbook()
    wb.remove(wb.active)
    usados = set()

    # Hoja de información (título, fecha sin hora y archivos ingresados).
    info = wb.create_sheet(_nombre_hoja("Información", usados))
    info["A1"] = titulo
    info["A1"].font = Font(bold=True, size=13)
    info["A2"] = "Generado: " + datetime.now().strftime("%Y-%m-%d")
    fila_info = 4
    if archivos:
        info[f"A{fila_info}"] = "Archivos ingresados:"
        info[f"A{fila_info}"].font = Font(bold=True)
        fila_info += 1
        for etiqueta, valor in archivos.items():
            if isinstance(valor, (list, tuple, set)):
                nombres = ", ".join(str(v) for v in valor) if valor else "—"
            else:
                nombres = str(valor) if valor else "—"
            info[f"A{fila_info}"] = f"• {etiqueta}: {nombres}"
            fila_info += 1
    info.column_dimensions["A"].width = 80

    secciones_validas = [
        (sub, df) for sub, df in secciones if isinstance(df, pd.DataFrame) and not df.empty
    ]
    if not secciones_validas:
        ws = wb.create_sheet(_nombre_hoja("Resultados", usados))
        ws["A1"] = "Sin resultados para mostrar."

    for subtitulo, df in secciones_validas:
        ws = wb.create_sheet(_nombre_hoja(subtitulo or "Resultados", usados))
        headers = list(df.columns)

        # Encabezado de la tabla.
        for col, header in enumerate(headers, start=1):
            celda = ws.cell(row=1, column=col, value=str(header))
            celda.fill = _HEADER_FILL
            celda.font = _HEADER_FONT
            celda.alignment = _WRAP_TOP

        # Una columna cuyo nombre empieza por "estado" colorea OK/diferencia.
        estado_idx = next(
            (i for i, h in enumerate(headers) if str(h).strip().lower().startswith("estado")),
            None,
        )

        for fila, (_, row) in enumerate(df.iterrows(), start=2):
            for col, header in enumerate(headers, start=1):
                texto, es_diff = _texto_celda(row[header], source_labels)
                celda = ws.cell(row=fila, column=col, value=texto)
                celda.alignment = _WRAP_TOP
                if es_diff:
                    celda.fill = _DIFF_FILL
                    celda.font = _DIFF_FONT
            if estado_idx is not None:
                est = ws.cell(row=fila, column=estado_idx + 1)
                es_ok = str(est.value).strip().lower() == "ok"
                est.fill = _OK_FILL if es_ok else _DIFF_FILL
                est.font = _OK_FONT if es_ok else _DIFF_FONT

        # Ancho de columnas aproximado al contenido del encabezado.
        for col, header in enumerate(headers, start=1):
            ws.column_dimensions[get_column_letter(col)].width = min(45, max(12, len(str(header)) + 4))
        ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()
