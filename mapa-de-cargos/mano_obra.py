"""Comparación de mano de obra: Informe de Costo vs registro de la ODS.

Ambos Excel traen la misma información bajo nombres de columna distintos. Este
módulo cruza a las personas por número de documento y, por cada campo
equivalente, indica si los valores coinciden o difieren.

Es la lógica de **producción** que consume la app web (``web_ui``). El notebook
``mano-obra.ipynb`` es la versión de exploración equivalente.

Modelo de datos del resultado (``comparar_mano_obra``):

- una fila por persona cruzada (presente en ambos archivos),
- columna ``Documento`` (solo dígitos),
- una columna por cada campo del ``MAPEO_COLUMNAS``, cuyo valor es una **lista**:
  ``[valor]`` si ambos coinciden, ``[valor_informe, valor_ods]`` si difieren,
- columna ``Estado revisión``: ``"ok"`` o ``"valores no coinciden: <campos>"``.

Las listas son la única fuente de verdad: una celda con dos elementos es una
inconsistencia, y eso es lo que la web y el Excel resaltan a nivel de celda.

Notas de formato del Informe:

- el encabezado real está en la fila 10 (índice 9),
- varias columnas con tildes llegan con el carácter corrupto en el xlsx
  (p. ej. ``Identificación``, ``Días Trabajados``), por eso las columnas que se
  usan se referencian **por posición** y se renombran a nombres limpios,
- la orden de servicio no es una columna directa: se extrae del texto de
  ``Nombre Centro Costo`` (``…Os050…`` -> ``50``).
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

INFORME_SHEET = "Informe"
INFORME_HEADER_ROW = 9  # el encabezado real está en la fila 10 (índice 9)

# Columnas del Informe usadas, por posición (los nombres del xlsx no son fiables).
INFORME_COLS_POR_POSICION = {
    0: "Tipo de pago",
    2: "Identificacion",
    4: "Nombres",
    5: "Apellidos",
    6: "Cargo",
    10: "Nombre Centro Costo",
    14: "Fecha de Ingreso",
    15: "Fecha de retiro",
    16: "Dias Trabajados",
}

# (etiqueta visible, columna en el Informe, columna en la ODS, tipo de comparación).
# tipo: "texto" (mayúsculas/sin acentos), "fecha" (por día) o "numero".
# Agrega aquí más campos a validar; el resto del flujo se adapta solo.
MAPEO_COLUMNAS = [
    ("OS", "OS", "No_de_orden_de_servicio_conocido_por_el_contratista", "numero"),
    ("Nombres", "Nombres", "Nombres", "texto"),
    ("Apellidos", "Apellidos", "Apellidos", "texto"),
    ("Cargo", "Cargo", "CargoContratoLaboral", "texto"),
    (
        "Fecha de Ingreso",
        "Fecha de Ingreso",
        "Fecha_de_inicio_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    (
        "Fecha de retiro",
        "Fecha de retiro",
        "Fecha_fin_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    ("Días Trabajados", "Dias Trabajados", "DiasTrabajadosEnMes", "numero"),
]

COL_DOC_ODS = "NumeroDocumento"
COL_DOCUMENTO = "Documento"
COL_ESTADO = "Estado revisión"
ESTADO_OK = "ok"


def solo_digitos(valor) -> str:
    """Deja solo los dígitos (para comparar documentos: ``91.499.442`` -> ``91499442``)."""
    return re.sub(r"\D", "", str(valor))


def norm_texto(valor) -> str:
    """Mayúsculas, sin acentos y con espacios colapsados, para comparar texto."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto)


def extraer_os(valor) -> int | None:
    """Extrae el número de orden de servicio del texto del centro de costo.

    ``"Ecop P&C Tabarca Grb Os050 - ..."`` -> ``50``.
    """
    if pd.isna(valor):
        return None
    match = re.search(r"OS\s*0*(\d+)", str(valor), re.IGNORECASE)
    return int(match.group(1)) if match else None


def _es_vacio(valor) -> bool:
    if valor is None:
        return True
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _presentacion(valor, tipo) -> str:
    """Valor 'bonito' que se muestra en la celda (fecha ISO, entero limpio, texto)."""
    if tipo == "fecha":
        fecha = pd.to_datetime(valor, errors="coerce")
        return "" if pd.isna(fecha) else fecha.strftime("%Y-%m-%d")
    if tipo == "numero":
        if _es_vacio(valor):
            return ""
        try:
            numero = float(str(valor).replace(",", "").strip())
            return str(int(numero)) if numero.is_integer() else str(numero)
        except ValueError:
            return re.sub(r"\s+", "", str(valor))
    # texto
    return "" if _es_vacio(valor) else str(valor)


def _comparable(valor, tipo) -> str:
    """Forma normalizada que decide si dos valores coinciden."""
    if tipo == "texto":
        return norm_texto(valor)
    return _presentacion(valor, tipo)  # fecha y numero ya quedan canónicos


def leer_informe(source) -> pd.DataFrame:
    """Lee la hoja 'Informe', renombra las columnas usadas y deriva la OS."""
    df = pd.read_excel(source, sheet_name=INFORME_SHEET, header=INFORME_HEADER_ROW)
    rename = {
        df.columns[idx]: nombre
        for idx, nombre in INFORME_COLS_POR_POSICION.items()
        if idx < len(df.columns)
    }
    df = df.rename(columns=rename)
    # Solo interesan las filas de Recobro (no las de Nómina).
    df = df[df["Tipo de pago"].map(norm_texto) == "RECOBRO"]
    df["OS"] = df["Nombre Centro Costo"].map(extraer_os) if "Nombre Centro Costo" in df.columns else None
    return df


def leer_ods(source) -> pd.DataFrame:
    """Lee el Excel de la ODS (una sola hoja con encabezados limpios)."""
    return pd.read_excel(source)


def comparar_mano_obra(informe_source, ods_source, mapeo=MAPEO_COLUMNAS) -> pd.DataFrame:
    """Cruza el Informe contra la ODS y devuelve el DataFrame de validación.

    ``informe_source`` / ``ods_source`` pueden ser rutas o buffers (lo que acepte
    ``pandas.read_excel``).
    """
    inf = leer_informe(informe_source)
    ods = leer_ods(ods_source)

    inf["_doc"] = inf["Identificacion"].map(solo_digitos)
    ods["_doc"] = ods[COL_DOC_ODS].map(solo_digitos)

    # Quitar filas sin documento y duplicados (cada lado puede repetir a la persona).
    inf = inf[inf["_doc"] != ""].drop_duplicates("_doc")
    ods = ods[ods["_doc"] != ""].drop_duplicates("_doc")
    ods_por_doc = ods.set_index("_doc")

    filas = []
    for _, registro in inf.iterrows():
        doc = registro["_doc"]
        if doc not in ods_por_doc.index:
            continue  # persona del Informe que no está en la ODS
        otro = ods_por_doc.loc[doc]

        fila = {COL_DOCUMENTO: doc}
        campos_diferentes = []
        for etiqueta, col_inf, col_ods, tipo in mapeo:
            val_inf = registro.get(col_inf)
            val_ods = otro.get(col_ods)
            disp_inf = _presentacion(val_inf, tipo)
            disp_ods = _presentacion(val_ods, tipo)
            if _comparable(val_inf, tipo) == _comparable(val_ods, tipo):
                fila[etiqueta] = [disp_inf]            # coinciden -> un solo elemento
            else:
                fila[etiqueta] = [disp_inf, disp_ods]  # difieren -> ambos valores
                campos_diferentes.append(etiqueta)

        fila[COL_ESTADO] = (
            ESTADO_OK
            if not campos_diferentes
            else "valores no coinciden: " + ", ".join(campos_diferentes)
        )
        filas.append(fila)

    columnas = [COL_DOCUMENTO] + [entrada[0] for entrada in mapeo] + [COL_ESTADO]
    return pd.DataFrame(filas, columns=columnas)
