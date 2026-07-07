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
  ``[valor]`` si ambos coinciden, ``[valor_informe, valor_ods]`` si difieren.

Las listas son la única fuente de verdad: una celda con dos elementos es una
inconsistencia, y eso es lo que la web y el PDF resaltan a nivel de celda. (Ya
no se emite una columna de estado/observaciones: el resaltado por celda es el
único indicador de inconsistencia.)

Notas de formato del Informe:

- el encabezado real está en la fila 10 (índice 9),
- las columnas usadas se localizan **por nombre** (normalizado: mayúsculas, sin
  acentos y espacios colapsados, con alias por si el nombre cambia), porque el
  layout del Informe varía entre exportes mensuales (distinto número y orden de
  columnas). Si una columna requerida no existe en ese archivo, el campo queda
  **vacío** en vez de tomar por error otra columna en esa posición,
- la orden de servicio no es una columna directa: se extrae del texto de
  ``Nombre Centro Costo`` (``…Os050…`` -> ``50``).

Notas de comparación de fechas:

- las fechas de actividades de la ODS (inicio/fin del trabajador para el
  contrato comercial u orden de servicio) corresponden a la vigencia del
  **contrato**, por lo que se comparan contra ``Fecha Inicio`` y
  ``Fecha Vencimiento`` del Informe, **no** contra ``Fecha de Ingreso`` /
  ``Fecha de retiro`` (vínculo laboral, un concepto distinto).
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

INFORME_SHEET = "Informe"
INFORME_HEADER_ROW = 9  # el encabezado real está en la fila 10 (índice 9)

# Columnas del Informe usadas, localizadas por NOMBRE (no por posición, que varía
# entre exportes). Para cada nombre interno se listan los alias aceptados; se
# comparan normalizados (mayúsculas/sin acentos/espacios colapsados), así que
# basta con que coincida alguno. Si ninguno aparece, el campo queda ausente.
INFORME_COLUMNAS = {
    "Tipo de pago": ["Valor pagado en nomina", "Tipo de pago", "Tipo"],
    "Identificacion": ["Identificacion", "Identificación", "Documento", "Cedula"],
    "Nombres": ["Nombres"],
    "Apellidos": ["Apellidos"],
    "Cargo": ["Cargo"],
    "Fecha Inicio": ["Fecha Inicio"],                # vigencia del contrato (inicio)
    "Fecha Vencimiento": ["Fecha Vencimiento"],      # vigencia del contrato (fin)
    "Nombre Centro Costo": ["Nombre Centro Costo"],
    "Dias Trabajados": ["Dias Trabajados", "Días Trabajados"],
    "Salario Diario Contratado": ["Salario Diario Contratado"],  # salario diario pactado
}

# (etiqueta visible, columna en el Informe, columna en la ODS, tipo de comparación).
# tipo: "texto" (mayúsculas/sin acentos), "cargo" (como texto pero ignorando el
#       marcador de progresión "(PROGRE)"), "fecha" (por día), "numero" o "moneda"
#       ("moneda" compara el valor numérico normalizado y lo muestra en COP).
# Agrega aquí más campos a validar; el resto del flujo se adapta solo.
MAPEO_COLUMNAS = [
    ("OS", "OS", "No_de_orden_de_servicio_conocido_por_el_contratista", "numero"),
    ("Nombres", "Nombres", "Nombres", "texto"),
    ("Apellidos", "Apellidos", "Apellidos", "texto"),
    ("Cargo", "Cargo", "CargoContratoLaboral", "texto"),
    # Las fechas de actividades de la ODS son la vigencia del CONTRATO, así que
    # se comparan contra Fecha Inicio / Fecha Vencimiento del Informe (no contra
    # Fecha de Ingreso / Fecha de retiro, que son del vínculo laboral).
    (
        "Fecha Inicio",
        "Fecha Inicio",
        "Fecha_de_inicio_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    (
        "Fecha Vencimiento",
        "Fecha Vencimiento",
        "Fecha_fin_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    ("Días Trabajados", "Dias Trabajados", "DiasTrabajadosEnMes", "numero"),
    # Salario diario: Informe ("Salario Diario Contratado") vs ODS
    # ("SalarioDiarioPesos"). Se compara como valor numérico normalizado y se
    # presenta formateado en pesos colombianos.
    ("Salario", "Salario Diario Contratado", "SalarioDiarioPesos", "moneda"),
]

COL_DOC_ODS = "NumeroDocumento"
COL_DOCUMENTO = "Documento"

# ---------------------------------------------------------------------------
# Formato ITALCO
# ---------------------------------------------------------------------------
# En ITALCO el Informe es la "progresión" mensual, con un layout distinto al de
# TABARCA:
#
# - una sola hoja cuyo nombre incluye el mes (``PROGRESION JULIO 2025``), así que
#   se lee la primera hoja, no una llamada "Informe";
# - el encabezado real no es la primera fila (hay filas de título arriba); se
#   detecta como la primera fila que trae a la vez ``Documento`` y
#   ``Nombre Completo``;
# - la primera columna no tiene nombre y marca ``ACTUAL`` / ``ANTERIOR`` /
#   ``DIFERENCIA`` por persona. Solo interesan las filas ``DIFERENCIA`` (el
#   ajuste del mes), así que se filtra por esa columna sin nombre;
# - el nombre viene como un único ``Nombre Completo`` (no separado en
#   Nombres/Apellidos), por lo que en la ODS se compara contra la combinación
#   ``Nombres + Apellidos`` (ver ``COL_NOMBRE_COMPLETO_ODS``);
# - las fechas comparadas son ``Fecha de Inicio`` y ``Fecha retiro`` del Informe
#   contra las fechas de actividades del contrato en la ODS.

# Columnas del Informe ITALCO, localizadas por NOMBRE (con alias). ``Documento``
# se renombra a ``Identificacion`` para reutilizar el mismo cruce por documento
# que TABARCA.
INFORME_ITALCO_COLUMNAS = {
    "Identificacion": ["Documento"],
    "Nombre Completo": ["Nombre Completo"],
    "Cargo": ["Cargo"],
    "Fecha de Inicio": ["Fecha de Inicio"],
    "Fecha retiro": ["Fecha retiro", "Fecha Retiro"],
}

# Columna derivada en la ODS: nombre completo = Nombres + Apellidos, para
# compararlo contra el único ``Nombre Completo`` del Informe ITALCO.
COL_NOMBRE_COMPLETO_ODS = "_NombreCompletoODS"

# Días del mes usados para pasar el Sueldo Base (mensual) a salario diario en
# ITALCO. La convención de nómina colombiana usa 30 días/mes, y se verificó que
# ``Sueldo Base (fila ACTUAL) / 30`` == ``SalarioDiarioPesos`` de la ODS.
DIAS_MES_ITALCO = 30

# Mapeo de comparación ITALCO (etiqueta, col Informe, col ODS, tipo). A
# diferencia de TABARCA, aquí las fechas de actividades de la ODS se comparan
# contra ``Fecha de Inicio`` / ``Fecha retiro`` del Informe (así lo define el
# formato ITALCO de la progresión).
MAPEO_ITALCO = [
    # "os": la OS del Informe sale del "Perfil Contable" (BCA OS 37 CONVENCIONAL)
    # y la de la ODS viene como "0DS37"; se comparan por su número (37).
    ("OS", "OS", "No_de_orden_de_servicio_conocido_por_el_contratista", "os"),
    ("Nombre Completo", "Nombre Completo", COL_NOMBRE_COMPLETO_ODS, "texto"),
    # "cargo": ignora el marcador de progresión "(PROGRE)" que el Informe añade y
    # la ODS no, para no marcar como diferencia cargos sustancialmente iguales.
    ("Cargo", "Cargo", "CargoContratoLaboral", "cargo"),
    (
        "Fecha Inicio",
        "Fecha de Inicio",
        "Fecha_de_inicio_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    (
        "Fecha Retiro",
        "Fecha retiro",
        "Fecha_fin_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio",
        "fecha",
    ),
    # Salario diario del Informe (Sueldo Base de la fila ACTUAL / 30, derivado en
    # ``leer_informe_italco``) contra el SalarioDiarioPesos de la ODS, en COP.
    ("Salario", "SalarioDiario", "SalarioDiarioPesos", "moneda"),
]


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


def norm_cargo(valor) -> str:
    """Normaliza un cargo para comparar Informe ITALCO contra ODS.

    El Informe de la progresión añade al cargo el marcador ``(PROGRE)`` /
    ``(PROGRESION)`` que la ODS no trae, así que ``AYUDANTE TECNICO A / TUBERIA
    C6 (PROGRE)`` y ``AYUDANTE TECNICO A / TUBERIA C6`` son el mismo cargo. Se
    quita ese paréntesis de progresión (además de la normalización de texto
    habitual) para que no cuente como diferencia; las diferencias reales del
    cargo (p. ej. ``E12`` vs ``E11``) se conservan.
    """
    texto = norm_texto(valor)                      # mayúsculas, sin acentos, espacios colapsados
    texto = re.sub(r"\(\s*PROG[A-Z]*\s*\)", "", texto)  # (PROGRE), (PROGRESION), (PROG)
    return re.sub(r"\s+", " ", texto).strip()


def extraer_os(valor) -> int | None:
    """Extrae el número de orden de servicio del texto del centro de costo.

    ``"Ecop P&C Tabarca Grb Os050 - ..."`` -> ``50``.
    """
    if pd.isna(valor):
        return None
    match = re.search(r"OS\s*0*(\d+)", str(valor), re.IGNORECASE)
    return int(match.group(1)) if match else None


def _os_comparable(valor) -> str:
    """Número de orden de servicio (último grupo de dígitos) para comparar.

    En ITALCO la OS llega con formatos distintos a cada lado: en el Informe como
    ``BCA OS 37 CONVENCIONAL`` y en la ODS como ``0DS37``. Se reduce ambos al
    número final (``37``) para compararlos sin depender del texto. Vacío si no
    hay dígitos.
    """
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    numeros = re.findall(r"\d+", str(valor))
    return str(int(numeros[-1])) if numeros else ""


def normalizar_moneda(valor) -> float | None:
    """Normaliza un valor monetario heterogéneo a ``float`` (o ``None``).

    Reutilizable para comparar y formatear cualquier salario/importe. Tolera:

    - ``None`` / ``NaN`` / cadenas vacías -> ``None``;
    - valores ya numéricos (``int`` / ``float``);
    - símbolo ``$``, espacios y demás caracteres no numéricos;
    - separadores de miles/decimales en convención colombiana
      (``1.234.567,89``) o anglosajona (``1,234,567.89``).

    Así ``"$ 120.000"``, ``"120000"`` y ``120000.0`` resultan en el mismo
    número y la comparación nunca depende de la forma del texto.
    """
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass

    # Si ya es numérico (no bool), usarlo directamente.
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    # Conservar solo dígitos y separadores; descarta "$", espacios, letras, etc.
    texto = re.sub(r"[^\d,.\-]", "", texto)
    if not texto or texto in {"-", ".", ",", "-.", "-,"}:
        return None

    # Resolver qué separador es el decimal y cuál el de miles.
    if "," in texto and "." in texto:
        # El separador que aparece más a la derecha es el decimal.
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        partes = texto.split(",")
        # ",XX" (1-2 dígitos finales) -> decimal; si no, separador de miles.
        if len(partes) == 2 and 1 <= len(partes[-1]) <= 2:
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "." in texto:
        partes = texto.split(".")
        if len(partes) > 2:
            # Varios puntos -> todos son de miles (1.234.567).
            texto = texto.replace(".", "")
        elif len(partes[-1]) == 3:
            # Un punto con 3 dígitos finales -> miles (120.000), no decimal.
            texto = texto.replace(".", "")
        # En otro caso (p. ej. "120.5") se deja como decimal.

    try:
        return float(texto)
    except ValueError:
        return None


def formatear_cop(valor) -> str:
    """Formatea un valor en moneda colombiana (COP): ``$120.000``.

    - separador de miles con punto, sin notación científica;
    - sin decimales cuando el valor es entero; con dos decimales (coma) si los
      tiene (``$120.000,50``);
    - cadena vacía cuando el valor no es interpretable como número.
    """
    numero = normalizar_moneda(valor)
    if numero is None:
        return ""

    if float(numero).is_integer():
        entero = int(round(numero))
        # f"{n:,}" usa coma de miles; se intercambia por punto (formato COP).
        return "$" + f"{entero:,}".replace(",", ".")

    # Dos decimales: 1,234,567.89 -> 1.234.567,89
    texto = f"{numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "$" + texto


def _es_vacio(valor) -> bool:
    if valor is None:
        return True
    try:
        return bool(pd.isna(valor))
    except (TypeError, ValueError):
        return False


def _presentacion(valor, tipo) -> str:
    """Valor 'bonito' que se muestra en la celda (fecha ISO, entero, texto, COP)."""
    if tipo == "fecha":
        # Normaliza a fecha por día (ignora la hora). Las fechas de estos Excel
        # llegan como Timestamp; se comparan por su parte de fecha en ISO
        # (``YYYY-MM-DD``) para que ``2025-06-08 00:00:00`` y ``2025-06-08``
        # coincidan. No se fuerza ``dayfirst`` porque rompería las cadenas ISO
        # (``2025-06-01`` se leería como 6 de enero). Nulos/vacíos/ inválidos
        # -> "" (así dos fechas ausentes se consideran iguales, no difieren).
        fecha = pd.to_datetime(valor, errors="coerce")
        return "" if pd.isna(fecha) else fecha.strftime("%Y-%m-%d")
    if tipo == "moneda":
        return formatear_cop(valor)
    if tipo == "os":
        return _os_comparable(valor)
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
    """Forma normalizada que decide si dos valores coinciden.

    Nunca compara cadenas crudas: cada tipo se reduce a una forma canónica
    (texto sin acentos, fecha por día, número entero, o importe numérico) para
    evitar falsos positivos por formato (``$120.000`` vs ``120000``, etc.).
    """
    if tipo == "texto":
        return norm_texto(valor)
    if tipo == "cargo":
        return norm_cargo(valor)
    if tipo == "os":
        return _os_comparable(valor)
    if tipo == "moneda":
        numero = normalizar_moneda(valor)
        return "" if numero is None else f"{numero:.2f}"
    return _presentacion(valor, tipo)  # fecha y numero ya quedan canónicos


def _partes_cargo(valor) -> set[str]:
    """Alternativas de un cargo separadas por ``/`` (normalizadas, sin ``(PROGRE)``).

    El Informe ITALCO expresa el cargo como alternativas separadas por ``/``
    (``ANDAMIERO B / D8``): cada tramo es una descripción válida del mismo cargo.
    """
    texto = norm_cargo(valor)
    if not texto:
        return set()
    return {parte.strip() for parte in texto.split("/") if parte.strip()}


def _coinciden(val_inf, val_ods, tipo) -> bool:
    """Decide si el valor del Informe y el de la ODS se consideran iguales.

    Para casi todos los tipos es la igualdad de su forma canónica. El tipo
    ``cargo`` es especial: como el Informe da alternativas separadas por ``/``
    (``ANDAMIERO B / D8``), basta con que **una** de ellas coincida con alguna de
    las de la ODS para tomarlo como el mismo cargo.
    """
    if tipo == "cargo":
        partes_inf = _partes_cargo(val_inf)
        partes_ods = _partes_cargo(val_ods)
        if not partes_inf or not partes_ods:
            return partes_inf == partes_ods  # ambos vacíos -> iguales; uno vacío -> difieren
        return bool(partes_inf & partes_ods)
    return _comparable(val_inf, tipo) == _comparable(val_ods, tipo)


def leer_informe(source) -> pd.DataFrame:
    """Lee la hoja 'Informe', localiza las columnas usadas por nombre y deriva la OS.

    Las columnas se buscan por nombre normalizado (con alias), no por posición,
    porque el layout varía entre exportes mensuales. Una columna que no aparezca
    simplemente no se renombra y queda ausente del resultado.
    """
    df = pd.read_excel(source, sheet_name=INFORME_SHEET, header=INFORME_HEADER_ROW)

    # nombre_normalizado -> nombre real (primera aparición gana).
    por_norm: dict[str, object] = {}
    for col in df.columns:
        por_norm.setdefault(norm_texto(col), col)

    rename = {}
    for destino, alias in INFORME_COLUMNAS.items():
        for nombre in alias:
            real = por_norm.get(norm_texto(nombre))
            if real is not None:
                rename[real] = destino
                break
    df = df.rename(columns=rename)

    # Solo interesan las filas de Recobro (no las de Nómina), si la columna existe.
    if "Tipo de pago" in df.columns:
        df = df[df["Tipo de pago"].map(norm_texto) == "RECOBRO"]
    df["OS"] = df["Nombre Centro Costo"].map(extraer_os) if "Nombre Centro Costo" in df.columns else None
    return df


def leer_ods(source) -> pd.DataFrame:
    """Lee el Excel de la ODS (una sola hoja con encabezados limpios)."""
    return pd.read_excel(source)


def leer_informe_italco(source) -> pd.DataFrame:
    """Lee el Informe ITALCO (progresión) y deja solo las filas de DIFERENCIA.

    Detecta la fila de encabezado real (la primera con ``Documento`` y
    ``Nombre Completo``, porque hay filas de título arriba) y localiza las columnas
    usadas por nombre (renombrando ``Documento`` -> ``Identificacion`` para
    reutilizar el cruce por documento).

    La primera columna, sin nombre, marca ``ACTUAL`` / ``ANTERIOR`` / ``DIFERENCIA``.
    El resultado son las filas ``DIFERENCIA`` (el ajuste del mes), pero antes de
    filtrar se guarda el ``Sueldo Base`` de la fila ``ACTUAL`` de cada persona para
    derivar el salario diario (``Sueldo Base / 30``), ya que en la fila DIFERENCIA
    la columna de salario es un delta, no el salario real.
    """
    crudo = pd.read_excel(source, sheet_name=0, header=None)
    fila_encabezado = 0
    for i in range(min(30, len(crudo))):
        celdas = {norm_texto(v) for v in crudo.iloc[i].tolist()}
        if "DOCUMENTO" in celdas and "NOMBRE COMPLETO" in celdas:
            fila_encabezado = i
            break
    df = pd.read_excel(source, sheet_name=0, header=fila_encabezado)

    # La primera columna no tiene nombre (pandas la llama "Unnamed: 0") y marca
    # ACTUAL / ANTERIOR / DIFERENCIA. Se conserva para separar esas filas.
    col_marca = df.columns[0]

    # nombre_normalizado -> nombre real (primera aparición gana).
    por_norm: dict[str, object] = {}
    for col in df.columns:
        por_norm.setdefault(norm_texto(col), col)

    rename = {}
    for destino, alias in INFORME_ITALCO_COLUMNAS.items():
        for nombre in alias:
            real = por_norm.get(norm_texto(nombre))
            if real is not None:
                rename[real] = destino
                break
    df = df.rename(columns=rename)

    marca = df[col_marca].map(norm_texto)
    doc = df["Identificacion"].map(solo_digitos)

    # Sueldo Base de la fila ACTUAL por persona: es el salario mensual vigente; la
    # fila DIFERENCIA solo trae el delta del mes. doc -> Sueldo Base ACTUAL.
    col_sueldo = por_norm.get(norm_texto("Sueldo Base"))
    sueldo_actual: dict[str, float] = {}
    if col_sueldo is not None:
        act = marca == "ACTUAL"
        for d, val in zip(doc[act], df.loc[act, col_sueldo]):
            monto = normalizar_moneda(val)
            if d and monto is not None:
                sueldo_actual[d] = monto

    # Filtrar al ajuste (DIFERENCIA).
    df = df[marca == "DIFERENCIA"].copy()

    # La OS no es una columna directa: se extrae del "Perfil Contable"
    # (``BCA OS 37 CONVENCIONAL`` -> ``37``), igual que TABARCA la extrae del
    # centro de costo.
    df["OS"] = df["Perfil Contable"] if "Perfil Contable" in df.columns else None

    # Salario diario = Sueldo Base ACTUAL / 30 (mensual -> diario), para comparar
    # contra el SalarioDiarioPesos de la ODS.
    doc_dif = df["Identificacion"].map(solo_digitos)
    df["SalarioDiario"] = [
        (sueldo_actual[d] / DIAS_MES_ITALCO) if d in sueldo_actual else None
        for d in doc_dif
    ]
    return df


def leer_ods_italco(source) -> pd.DataFrame:
    """Lee la ODS y agrega el nombre completo (Nombres + Apellidos).

    El Informe ITALCO trae un único ``Nombre Completo``; la ODS lo tiene separado
    en ``Nombres`` y ``Apellidos``, así que se combinan en ``COL_NOMBRE_COMPLETO_ODS``
    para poder compararlos.
    """
    df = leer_ods(source)
    nombres = df["Nombres"].fillna("").astype(str).str.strip() if "Nombres" in df.columns else ""
    apellidos = df["Apellidos"].fillna("").astype(str).str.strip() if "Apellidos" in df.columns else ""
    df[COL_NOMBRE_COMPLETO_ODS] = (nombres + " " + apellidos).str.strip()
    return df


def _leer_y_concatenar(sources, lector) -> pd.DataFrame:
    """Lee una o varias fuentes (ruta/buffer) con ``lector`` y las concatena.

    Permite pasar más de un Excel por lado: cada archivo se lee por separado y
    el resultado se une en un único DataFrame antes del cruce.
    """
    if not isinstance(sources, (list, tuple)):
        sources = [sources]
    partes = [lector(src) for src in sources]
    if not partes:
        return pd.DataFrame()
    return pd.concat(partes, ignore_index=True)


def comparar_mano_obra(informe_source, ods_source, mapeo=None, formato="tabarca") -> pd.DataFrame:
    """Cruza el Informe contra la ODS y devuelve el DataFrame de validación.

    ``informe_source`` / ``ods_source`` pueden ser una ruta/buffer (lo que acepte
    ``pandas.read_excel``) o una **lista** de varias fuentes; en ese caso cada
    archivo se lee por separado y se concatena antes de cruzar, de modo que una
    persona de cualquier Informe puede emparejarse con cualquier ODS.

    ``formato`` (``tabarca`` / ``italco``) elige el layout del Informe y el mapeo
    de comparación: TABARCA usa la hoja ``Informe`` (fila de recobro, OS, salario…)
    y ITALCO la progresión (filas de DIFERENCIA, nombre completo y fechas de
    actividades). ``mapeo`` puede forzarse; por defecto se toma el del formato.

    En **ambos** formatos se listan todas las personas del Informe aunque no estén
    en la ODS: las que no cruzan aparecen con el lado ODS en blanco (todas sus
    celdas quedan resaltadas), para ver de un vistazo quién está en el Informe pero
    falta en la ODS.
    """
    es_italco = str(formato).lower() == "italco"
    if es_italco:
        inf = _leer_y_concatenar(informe_source, leer_informe_italco)
        ods = _leer_y_concatenar(ods_source, leer_ods_italco)
        if mapeo is None:
            mapeo = MAPEO_ITALCO
    else:
        inf = _leer_y_concatenar(informe_source, leer_informe)
        ods = _leer_y_concatenar(ods_source, leer_ods)
        if mapeo is None:
            mapeo = MAPEO_COLUMNAS

    inf["_doc"] = inf["Identificacion"].map(solo_digitos)
    ods["_doc"] = ods[COL_DOC_ODS].map(solo_digitos)

    # Quitar filas sin documento y duplicados (cada lado puede repetir a la persona).
    inf = inf[inf["_doc"] != ""].drop_duplicates("_doc")
    ods = ods[ods["_doc"] != ""].drop_duplicates("_doc")
    ods_por_doc = ods.set_index("_doc")

    filas = []
    for _, registro in inf.iterrows():
        doc = registro["_doc"]
        en_ods = doc in ods_por_doc.index
        # Se lista a toda persona del Informe; si no cruza, el lado ODS queda
        # vacío (None) y sus celdas quedan resaltadas como faltantes.
        otro = ods_por_doc.loc[doc] if en_ods else None

        fila = {COL_DOCUMENTO: doc}
        for etiqueta, col_inf, col_ods, tipo in mapeo:
            val_inf = registro.get(col_inf)
            val_ods = otro.get(col_ods) if otro is not None else None
            disp_inf = _presentacion(val_inf, tipo)
            disp_ods = _presentacion(val_ods, tipo)
            if _coinciden(val_inf, val_ods, tipo):
                fila[etiqueta] = [disp_inf]            # coinciden -> un solo elemento
            else:
                fila[etiqueta] = [disp_inf, disp_ods]  # difieren -> ambos valores
        filas.append(fila)

    # Sin columna de estado/observaciones: el resaltado por celda (listas de dos
    # elementos) es el único indicador de inconsistencia.
    columnas = [COL_DOCUMENTO] + [entrada[0] for entrada in mapeo]
    return pd.DataFrame(filas, columns=columnas)
