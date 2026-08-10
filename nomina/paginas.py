"""Iteracion de paginas de un PDF liberando su cache al avanzar.

Gemelo de ``facturacion/paginas.py``: se duplica para que cada paquete siga
siendo autocontenido, igual que ``depuracion.py``.
"""
from __future__ import annotations


def iter_paginas(pdf):
    """Itera ``pdf.pages`` liberando el cache de cada pagina al pasar a la siguiente.

    ``pdfplumber`` cachea por pagina los objetos que parsea (``_objects``,
    ``_layout``, ``_edges``, ``_rect_edges``, ``_curve_edges``) y el textmap, y
    **no los libera** al avanzar: ``pdf.pages`` retiene cada pagina ya visitada,
    asi que la memoria crece de forma acumulativa a lo largo del documento (se
    midio ~1.6 MB por pagina en una planilla densa, es decir ~200 MB en 120
    paginas) y solo se recupera al cerrar el PDF completo.

    ``page.close()`` hace ``flush_cache()`` mas ``get_textmap.cache_clear()``.
    Esas propiedades son **perezosas**: si algo vuelve a pedirlas se recalculan
    solas, por lo que liberarlas no altera ningun resultado, solo el momento en
    que se hace el trabajo.

    Se usa un generador con ``try/finally`` en lugar de un ``page.close()`` al
    final del bucle porque varios bucles hacen ``continue`` a nivel de pagina;
    el ``finally`` corre igual en esos caminos y tambien ante ``break`` o
    excepcion, cuando el generador se cierra.
    """
    for page in pdf.pages:
        try:
            yield page
        finally:
            page.close()
