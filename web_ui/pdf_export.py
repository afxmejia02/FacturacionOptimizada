"""Build a multi-page PDF with the review results for download.

Edge's "print to PDF" only captures the visible page; this module renders the
full result set (every row, across as many pages as needed) plus the names of
the uploaded files, using reportlab. All styling stays here so the Streamlit
entry point only wires the download button.
"""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

_OK_BG = colors.HexColor("#D4EDDA")
_OK_FG = colors.HexColor("#155724")
_DIFF_BG = colors.HexColor("#F8D7DA")
_DIFF_FG = colors.HexColor("#721C24")
_HEADER_BG = colors.HexColor("#1F4E78")
_GRID = colors.HexColor("#BBBBBB")

_PAGE = landscape(A4)
_MARGIN = 1.0 * cm


def _build_styles():
    base = getSampleStyleSheet()
    titulo = ParagraphStyle("titulo", parent=base["Title"], fontSize=15, spaceAfter=6)
    meta = ParagraphStyle("meta", parent=base["BodyText"], fontSize=9, leading=12)
    seccion = ParagraphStyle("seccion", parent=base["Heading2"], fontSize=11, spaceBefore=10, spaceAfter=4)
    cell = ParagraphStyle("cell", parent=base["BodyText"], fontSize=7, leading=8.5, wordWrap="CJK")
    head = ParagraphStyle(
        "head", parent=cell, fontName="Helvetica-Bold", textColor=colors.white
    )
    return {"titulo": titulo, "meta": meta, "seccion": seccion, "cell": cell, "head": head}


def _texto_celda(value, source_labels):
    """Return ``(html_text, es_diferencia)`` for a single DataFrame cell.

    List cells (mano de obra) hold one value when both fuentes coinciden, or two
    when difieren; in that case both se muestran y la celda se marca distinta.
    """
    if isinstance(value, (list, tuple)):
        items = ["" if v is None else str(v) for v in value]
        if len(items) <= 1:
            return (items[0] if items else ""), False
        etiquetas = source_labels or ("", "")
        izquierda = f"<b>{etiquetas[0]}:</b> {items[0]}" if etiquetas[0] else items[0]
        derecha = f"<b>{etiquetas[1]}:</b> {items[1]}" if etiquetas[1] else items[1]
        return f"{izquierda}<br/>{derecha}", True

    if value is None:
        return "", False
    try:
        if pd.isna(value):
            return "", False
    except (TypeError, ValueError):
        pass
    return str(value), False


def _tabla_desde_df(df, source_labels, styles, ancho_disponible):
    headers = list(df.columns)
    estado_idx = next(
        (i for i, h in enumerate(headers) if str(h).strip().lower().startswith("estado")),
        None,
    )
    tiene_listas = any(
        df[h].map(lambda v: isinstance(v, (list, tuple))).any() for h in headers
    )

    filas = [[Paragraph(str(h), styles["head"]) for h in headers]]
    comandos = [
        ("BACKGROUND", (0, 0), (-1, 0), _HEADER_BG),
        ("GRID", (0, 0), (-1, -1), 0.4, _GRID),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]

    for r, (_, row) in enumerate(df.iterrows(), start=1):
        celdas = []
        cols_diferentes = []
        for c, h in enumerate(headers):
            texto, es_diff = _texto_celda(row[h], source_labels)
            celdas.append(Paragraph(texto, styles["cell"]))
            if es_diff:
                cols_diferentes.append(c)
        filas.append(celdas)

        estado_ok = None
        if estado_idx is not None:
            estado_ok = str(row[headers[estado_idx]]).strip().lower() == "ok"

        if tiene_listas:
            # Colorear solo las celdas inconsistentes + la celda de estado.
            for c in cols_diferentes:
                comandos.append(("BACKGROUND", (c, r), (c, r), _DIFF_BG))
                comandos.append(("TEXTCOLOR", (c, r), (c, r), _DIFF_FG))
            if estado_idx is not None:
                bg, fg = (_OK_BG, _OK_FG) if estado_ok else (_DIFF_BG, _DIFF_FG)
                comandos.append(("BACKGROUND", (estado_idx, r), (estado_idx, r), bg))
                comandos.append(("TEXTCOLOR", (estado_idx, r), (estado_idx, r), fg))
        elif estado_idx is not None:
            # Colorear toda la fila según el estado.
            bg, fg = (_OK_BG, _OK_FG) if estado_ok else (_DIFF_BG, _DIFF_FG)
            comandos.append(("BACKGROUND", (0, r), (-1, r), bg))
            comandos.append(("TEXTCOLOR", (0, r), (-1, r), fg))

    ancho_col = ancho_disponible / max(len(headers), 1)
    tabla = Table(filas, colWidths=[ancho_col] * len(headers), repeatRows=1)
    tabla.setStyle(TableStyle(comandos))
    return tabla


def build_results_pdf(titulo, archivos, secciones, source_labels=None):
    """Render the results PDF and return its bytes.

    - ``titulo``: encabezado del documento.
    - ``archivos``: dict ``{"PDF": "...", "Excel": "..."}`` o con listas de
      nombres; se listan como "Archivos ingresados".
    - ``secciones``: lista de ``(subtitulo, DataFrame)``; cada DataFrame se
      pinta como tabla (paginada automáticamente).
    - ``source_labels``: par ``("Informe", "ODS")`` para las celdas con dos
      valores (mano de obra); ``None`` para el resto.
    """
    styles = _build_styles()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=_PAGE,
        leftMargin=_MARGIN,
        rightMargin=_MARGIN,
        topMargin=_MARGIN,
        bottomMargin=_MARGIN,
        title=titulo,
    )
    ancho_disponible = _PAGE[0] - 2 * _MARGIN

    elementos = [Paragraph(titulo, styles["titulo"])]
    elementos.append(
        Paragraph(
            "Generado: " + datetime.now().strftime("%Y-%m-%d %H:%M"), styles["meta"]
        )
    )

    if archivos:
        lineas = ["<b>Archivos ingresados:</b>"]
        for etiqueta, valor in archivos.items():
            if isinstance(valor, (list, tuple, set)):
                nombres = ", ".join(str(v) for v in valor) if valor else "—"
            else:
                nombres = str(valor) if valor else "—"
            lineas.append(f"• {etiqueta}: {nombres}")
        elementos.append(Paragraph("<br/>".join(lineas), styles["meta"]))

    elementos.append(Spacer(1, 0.3 * cm))

    secciones_validas = [
        (sub, df) for sub, df in secciones if isinstance(df, pd.DataFrame) and not df.empty
    ]
    if not secciones_validas:
        elementos.append(Paragraph("Sin resultados para mostrar.", styles["meta"]))

    for subtitulo, df in secciones_validas:
        if subtitulo:
            elementos.append(Paragraph(subtitulo, styles["seccion"]))
        elementos.append(_tabla_desde_df(df, source_labels, styles, ancho_disponible))
        elementos.append(Spacer(1, 0.25 * cm))

    doc.build(elementos)
    buffer.seek(0)
    return buffer.getvalue()
