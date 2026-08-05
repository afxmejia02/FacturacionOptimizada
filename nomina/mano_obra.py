"""Mano de obra: cruza el Informe de Costo contra el registro de la ODS.

Los dos Excel traen la misma informacion con nombres de columna distintos. Se
cruza por documento y, por cada campo equivalente, la celda resultante es una
lista: ``[valor]`` si coinciden, ``[informe, ods]`` si difieren (eso es lo que
la web resalta).

Las decisiones de formato (por que las columnas se buscan por nombre, contra
que fechas se compara, que es la "progresion" de ITALCO) estan en README.md.
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

INFORME_SHEET = "Informe"
INFORME_HEADER_ROW = 9  # respaldo si no se detecta el encabezado (ver _detectar_fila_encabezado_informe)

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
    ("Salario", "Salario Diario Contratado", "SalarioDiarioPesos", "moneda"),
]

COL_DOC_ODS = "NumeroDocumento"
COL_DOCUMENTO = "Documento"

# --- Formato ITALCO (la "progresion" mensual). Ver README.md. ---------------

# ``Documento`` se renombra a ``Identificacion`` para reutilizar el cruce de TABARCA.
INFORME_ITALCO_COLUMNAS = {
    "Identificacion": ["Documento"],
    "Nombre Completo": ["Nombre Completo"],
    "Cargo": ["Cargo"],
    "Fecha de Inicio": ["Fecha de Inicio"],
    "Fecha retiro": ["Fecha retiro", "Fecha Retiro"],
}

# Nombre completo derivado en la ODS (Nombres + Apellidos).
COL_NOMBRE_COMPLETO_ODS = "_NombreCompletoODS"

DIAS_MES_ITALCO = 30  # convencion de nomina colombiana: mensual / 30 = diario

MAPEO_ITALCO = [
    ("OS", "OS", "No_de_orden_de_servicio_conocido_por_el_contratista", "os"),
    ("Nombre Completo", "Nombre Completo", COL_NOMBRE_COMPLETO_ODS, "texto"),
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
    ("Salario", "SalarioDiario", "SalarioDiarioPesos", "moneda"),
]


def solo_digitos(valor) -> str:
    """Deja solo los digitos: ``91.499.442`` -> ``91499442``.

    Un float entero se pasa antes a ``int``: Excel puede traer el documento como
    numero y ``str(1096198448.0)`` dejaria un ``0`` de mas al final.
    """
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\D", "", str(valor))


def norm_texto(valor) -> str:
    """Mayúsculas, sin acentos y con espacios colapsados, para comparar texto."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto)


def norm_cargo(valor) -> str:
    """Normaliza un cargo, ignorando el marcador ``(PROGRE)`` de la progresion.

    Asi ``AYUDANTE TECNICO A C6 (PROGRE)`` y ``AYUDANTE TECNICO A C6`` son el
    mismo cargo, pero ``E12`` vs ``E11`` sigue siendo una diferencia.
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
    """Numero de la orden de servicio (ultimo grupo de digitos).

    Reduce ``BCA OS 37 CONVENCIONAL``, ``37`` y ``0DS37`` a ``"37"``. Un float
    entero se pasa a ``int``: de ``37.0`` se tomaria el ``0`` final.
    """
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    numeros = re.findall(r"\d+", str(valor))
    return str(int(numeros[-1])) if numeros else ""


def normalizar_moneda(valor) -> float | None:
    """Normaliza un importe heterogeneo a ``float`` (o ``None``).

    Tolera nulos, numeros, ``$``, y separadores en convencion colombiana
    (``1.234.567,89``) o anglosajona (``1,234,567.89``).
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
    """Formatea en pesos colombianos: ``120000`` -> ``$120.000``.

    Sin decimales si el valor es entero; con coma decimal si los tiene. Cadena
    vacia si no es interpretable como numero.
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
    """Decide si el valor del Informe y el de la ODS son iguales.

    Es la igualdad de su forma canonica, salvo ``cargo``: el Informe da
    alternativas separadas por ``/`` y basta que una coincida.
    """
    if tipo == "cargo":
        partes_inf = _partes_cargo(val_inf)
        partes_ods = _partes_cargo(val_ods)
        if not partes_inf or not partes_ods:
            return partes_inf == partes_ods  # ambos vacíos -> iguales; uno vacío -> difieren
        return bool(partes_inf & partes_ods)
    return _comparable(val_inf, tipo) == _comparable(val_ods, tipo)


def _detectar_fila_encabezado_informe(xls: pd.ExcelFile, max_filas: int = 25) -> int:
    """Fila del encabezado real de la hoja ``Informe``.

    Varia entre exportes mensuales, asi que se busca la primera fila que trae a
    la vez una columna de identificacion y ``Nombres``. Si no aparece, se usa
    ``INFORME_HEADER_ROW``.
    """
    crudo = xls.parse(INFORME_SHEET, header=None, nrows=max_filas)
    alias_doc = {norm_texto(a) for a in INFORME_COLUMNAS["Identificacion"]}
    alias_nombres = {norm_texto(a) for a in INFORME_COLUMNAS["Nombres"]}
    for i in range(len(crudo)):
        celdas = {norm_texto(v) for v in crudo.iloc[i].tolist()}
        if celdas & alias_doc and celdas & alias_nombres:
            return i
    return INFORME_HEADER_ROW


def leer_informe(source) -> pd.DataFrame:
    """Lee la hoja ``Informe``, localiza las columnas por nombre y deriva la OS.

    La fila del encabezado y los nombres de columna varian entre exportes; una
    columna que no aparezca queda ausente en vez de tomarse por posicion.
    """
    # Se cierra explícitamente (``with``): ``pd.ExcelFile`` mantiene el archivo
    # abierto hasta que se cierra o se recolecta la basura, y en Windows un
    # archivo con el handle abierto no se puede borrar (el ``TemporaryDirectory``
    # del llamador falla con "WinError 32" al limpiar si no se cierra aquí).
    with pd.ExcelFile(source) as xls:
        fila_encabezado = _detectar_fila_encabezado_informe(xls)
        df = xls.parse(INFORME_SHEET, header=fila_encabezado)

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
    """Lee el Informe ITALCO (progresion) y deja solo las filas de DIFERENCIA.

    Antes de filtrar guarda el ``Sueldo Base`` de la fila ``ACTUAL`` de cada
    persona: en la fila DIFERENCIA esa columna es un delta, no el salario real.
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

    # La OS varía entre exportes: unos la traen dentro del "Perfil Contable"
    # (``BCA OS 37 CONVENCIONAL``, igual que TABARCA la extrae del centro de
    # costo) y otros como columna propia ``OS`` con el número suelto (``37``).
    # Se acepta cualquiera de las dos; ``_os_comparable`` se queda con el número.
    # Se busca sobre ``por_norm`` (nombres originales) porque ninguna de las dos
    # entra en el rename, así que su nombre real no cambió.
    col_os = por_norm.get(norm_texto("Perfil Contable"))
    if col_os is None:
        col_os = por_norm.get(norm_texto("OS"))
    df["OS"] = df[col_os] if col_os is not None else None

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
    """Cruza el Informe contra la ODS y devuelve el DataFrame de validacion.

    Cada lado acepta una fuente o una lista (se concatenan antes de cruzar).
    ``formato`` elige el layout: ``tabarca`` (hoja ``Informe``) o ``italco``
    (progresion). Se listan todas las personas del Informe: las que no estan en
    la ODS salen con ese lado en blanco.
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
