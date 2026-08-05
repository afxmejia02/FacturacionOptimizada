"""Conciliacion de nomina: desprendibles contra transferencias y seguridad social.

Modulos:
  formato.py            limpieza de numeros, documentos y lineas de los PDF
  desprendibles.py      lectura de desprendibles (TABARCA / ITALCO)
  transferencias.py     lectura de transferencias bancarias
  seguridad_social.py   lectura de planillas de seguridad social (IBC)
  conciliacion.py       cruce de los tres origenes
  mano_obra.py          Informe de Costo contra el registro de la ODS
"""
from .conciliacion import conciliar
from .desprendibles import procesar_desprendibles
from .formato import formatear_valores
from .seguridad_social import procesar_seguridad_social
from .transferencias import match_linea_transferencia, procesar_transferencias

__all__ = [
    "conciliar",
    "formatear_valores",
    "match_linea_transferencia",
    "procesar_desprendibles",
    "procesar_seguridad_social",
    "procesar_transferencias",
]
