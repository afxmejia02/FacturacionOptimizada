"""Normalizacion de textos, fechas y cantidades del PDF y del Excel.

Cada nombre/fecha/cantidad se reduce a una forma canonica antes de comparar,
porque los dos archivos los escriben distinto. Ver README.md.
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

from .depuracion import DEBUG, debug as _debug


_RE_RECATEGORIZADO = re.compile(r"RECATEGORIZ\w*.*?\bCOMO\b.*?([A-Z]{1,3}\s*\d+)")
_RE_EF = re.compile(r"\bE\s*Y\s*F\b")
_RE_24H = re.compile(r"(?<!\d)24(?!\d)")

def _normalizar_texto_equipo(texto):
    if not isinstance(texto, str):
        return texto
    texto = texto.replace("\n", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()

def limpiar_nombre_equipo(texto):
    """Limpia el nombre de un equipo extraído de una celda del PDF.

    Algunos PDFs arrastran texto duplicado/superpuesto **después** del
    nombre real (p. ej. ``"MOTOSOLDADOR ... (24 H) Motoso"``: el ``Motoso``
    es un fragmento espurio —el inicio del propio nombre— de un texto que se
    superpone). Ese sobrante se descarta **solo si es un fragmento inicial
    duplicado del nombre**; una continuación legítima tras un paréntesis
    (p. ej. ``"Torno ... (Diurno / Nocturno) para bridas >4 NPS <= 48 NPS"``)
    **se conserva íntegra** para no suprimir información que sí está en el PDF
    y coincide con el Excel.
    """
    limpio = _normalizar_texto_equipo(texto)
    if not isinstance(limpio, str):
        return limpio
    idx = limpio.rfind(")")
    if idx != -1:
        base = limpio[: idx + 1].strip()
        cola = limpio[idx + 1:].strip()
        # Solo se recorta si la cola es un duplicado del inicio del nombre.
        cola_norm = clave_equipo(cola)
        if cola_norm and clave_equipo(base).startswith(cola_norm):
            limpio = base
    return limpio

def clave_equipo(texto):
    """Clave robusta para emparejar equipos/servicios entre PDF y Excel.

    Pliega acentos y mayúsculas, descarta comillas/paréntesis/comas y otros
    signos, elimina las conjunciones sueltas (y/o/e/u) y **descarta todos los
    espacios**. Así:

    - ``Camperos y camionetas ... (10 Horas)`` (PDF) empareja con
      ``Camperos o camionetas ... (10 Horas)`` (Excel);
    - ``... (10H)`` (PDF) empareja con ``... (10 H)`` (Excel): la diferencia
      de espacios deja de importar;

    pero sigue siendo distinto de la variante ``(24 Horas)`` porque conserva
    los dígitos.
    """
    if texto is None:
        return ""
    plano = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower()
    plano = re.sub(r"[^a-z0-9]+", " ", plano)
    # Quitar conjunciones sueltas (requiere espacios como límites de palabra)
    # ANTES de eliminar los espacios.
    plano = re.sub(r"\b[yoeu]\b", " ", plano)
    # El Excel a veces pega la unidad "DÍA"/"DÍAS" al final de la descripción
    # (p. ej. "... (24 Horas) DÍA") y el PDF no la trae; se descarta esa unidad
    # final para que emparejen. No se toca "horas": ahí sí discrimina la tarifa.
    plano = re.sub(r"\bdias?\b\s*$", " ", plano)
    # Sin espacios: "(10 H)" y "(10H)" producen la misma clave.
    return re.sub(r"\s+", "", plano)

def normalizar_perfil(valor):
    if not isinstance(valor, str):
        return valor
    texto = valor.strip()
    texto = texto.replace("Nivel", "").replace("Perfil", "")
    return texto.replace("/", "").strip()

def normalizar_busqueda(texto):
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto).strip().lower())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texto)

def normalizar_fecha(valor):
    meses = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "setiembre": 9, "octubre": 10,
        "noviembre": 11, "diciembre": 12,
    }

    if valor is None:
        return None

    texto = str(valor).strip().lower()
    fecha = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    if not pd.isna(fecha):
        return fecha.normalize()

    match = re.search(r"(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+de\s+(\d{4})", texto)
    if match:
        dia = int(match.group(1))
        mes = meses.get(match.group(2))
        anio = int(match.group(3))
        if mes is not None:
            try:
                return pd.Timestamp(anio, mes, dia)
            except ValueError:
                return None
    return None

def es_celda_vacia(valor):
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except Exception:
        pass

    texto_raw = str(valor)
    texto = unicodedata.normalize("NFKC", texto_raw)

    # Replace a broad set of invisible/whitespace characters with a single space
    texto = re.sub(r"[\s\u00A0\u2007\u202F\u200B\uFEFF\u2060\u200C\u200D]+", " ", texto)
    texto = texto.strip()

    # Remove surrounding matching quote characters repeatedly (handles '"   "')
    QUOTES = '"\'\u201C\u201D\u201E\u201F\u00AB\u00BB\u2039\u203A'
    while len(texto) >= 2 and texto[0] in QUOTES and texto[-1] in QUOTES:
        texto = texto[1:-1].strip()

    # Remove remaining quote characters and collapse interior whitespace
    texto = texto.replace('"', "").replace("'", "").replace("\u0000", "")
    texto = re.sub(r"\s+", "", texto).lower()

    if not texto:
        if DEBUG:
            _debug(f"_es_celda_vacia: raw={repr(texto_raw)} -> normalized empty string")
        return True
    if DEBUG:
        _debug(f"_es_celda_vacia: raw={repr(texto_raw)} -> normalized={repr(texto)}")
    return texto in {"nan", "none", "null"}

def parsear_cantidad(valor):
    """Convierte la cantidad de una planilla a ``float``.

    Convención **colombiana** (la que usan tanto el PDF como el Excel): el
    **punto es separador de miles** y la **coma es el separador decimal**.
    Ejemplos: ``3.139`` -> 3139, ``3.139,00`` -> 3139, ``1.452,6`` -> 1452.6,
    ``153,67`` -> 153.67, ``7,7`` -> 7.7.
    """
    if valor is None:
        return None

    # Primer token numérico (dígitos con . o , internos).
    match = re.search(r"\d[\d.,]*\d|\d", str(valor))
    if not match:
        return None
    # Punto = miles (se elimina); coma = decimal (se vuelve punto).
    token = match.group().replace(".", "").replace(",", ".")

    try:
        return float(token)
    except ValueError:
        return None

def parsear_observacion_perfil(observacion):
    """Interpreta la columna 'Observaciones' de la planilla de perfiles.

    Devuelve ``(recategorizado, es_ef, no_facturable, es_24h)``:

    - ``recategorizado`` – nivel al que se recategoriza (p. ej. ``"B4"``) o
      ``None`` si no hay recategorización.
    - ``es_ef`` – ``True`` si aparece el marcador **"E y F"**: el turno se
      cuenta como **1 unidad** aunque la jornada sea de 24 horas.
    - ``no_facturable`` – ``True`` si aparece **"NO FACTURABLE"**: no cuenta.
    - ``es_24h`` – ``True`` si la observación indica jornada de **24 horas**
      (``24``, ``24H``, ``24HRS``, ``24 HORAS``): el turno cuenta **1/3**.

    Los marcadores son independientes: pueden coexistir y en cualquier orden.
    """
    if observacion is None:
        return None, False, False, False
    texto = unicodedata.normalize("NFKD", str(observacion)).encode("ascii", "ignore").decode()
    texto = texto.replace('"', " ").replace("'", " ")
    texto = re.sub(r"\s+", " ", texto).upper().strip()
    if not texto:
        return None, False, False, False

    no_facturable = "NO FACTURABLE" in texto
    es_ef = _RE_EF.search(texto) is not None
    es_24h = _RE_24H.search(texto) is not None
    m = _RE_RECATEGORIZADO.search(texto)
    recategorizado = re.sub(r"\s+", "", m.group(1)) if m else None
    return recategorizado, es_ef, no_facturable, es_24h

def _buscar_indice_columna(row, opciones):
    opciones_norm = tuple(normalizar_busqueda(opcion) for opcion in opciones)
    for idx, cell in enumerate(row):
        cell_norm = normalizar_busqueda(cell)
        if any(opcion in cell_norm for opcion in opciones_norm):
            return idx
    return None
