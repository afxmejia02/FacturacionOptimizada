"""Validacion de perfiles, servicios y equipos: PDF contra el Excel historico.

Modulos:
  normalizacion.py  formas canonicas de textos, fechas y cantidades
  pdf.py            extraccion de conteos desde el informe PDF
  histograma.py     lectura del Excel historico
"""
from .histograma import (
    col_codigo_tarifa,
    leer_excel_facturacion,
    leer_histograma_largo,
    prefijos_seccion_pdf,
)
from .normalizacion import (
    clave_equipo,
    es_celda_vacia,
    limpiar_nombre_equipo,
    normalizar_busqueda,
    normalizar_fecha,
    normalizar_perfil,
    parsear_cantidad,
    parsear_observacion_perfil,
)
from .pdf import (
    extraer_conteo_pdf,
    extraer_perfiles_pdf,
    extraer_registros_etiqueta,
    extraer_valor_etiqueta,
)

__all__ = [
    "clave_equipo",
    "col_codigo_tarifa",
    "es_celda_vacia",
    "extraer_conteo_pdf",
    "extraer_perfiles_pdf",
    "extraer_registros_etiqueta",
    "extraer_valor_etiqueta",
    "leer_excel_facturacion",
    "leer_histograma_largo",
    "limpiar_nombre_equipo",
    "normalizar_busqueda",
    "normalizar_fecha",
    "normalizar_perfil",
    "parsear_cantidad",
    "parsear_observacion_perfil",
    "prefijos_seccion_pdf",
]
