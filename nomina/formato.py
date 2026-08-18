"""Limpieza y formato de numeros, documentos y lineas de texto de los PDF."""
from __future__ import annotations

import re
import unicodedata


def _limpiar_numero(valor):
    """
    Limpia y convierte una cadena numérica a float.
    
    Args:
        valor (str): Representación en texto del número
        
    Returns:
        float: Valor numérico limpio
    """
    if not valor: 
        return None
    valor = valor.replace("$", "").strip()
    if "," in valor and "." in valor:
        if valor.rfind(",") > valor.rfind("."):
            valor = valor.replace(".", "").replace(",", ".")
        else:
            valor = valor.replace(",", "")
    elif "," in valor:
        partes = valor.split(",")
        if len(partes[-1]) == 2:
            valor = valor.replace(".", "").replace(",", ".")
        else:
            valor = valor.replace(",", "")
    elif "." in valor:
        partes = valor.split(".")
        if len(partes[-1]) == 2:
            valor = valor.replace(",", "")
        elif valor.count(".") > 1:
            valor = valor.replace(".", "")
    valor = valor.strip()
    try:
        return float(valor)
    except:
        return None


def _limpiar_doc(col):
    """
    Limpia números de documento o identificación.
    
    Args:
        col (pd.Series): Serie de Pandas para limpiar
        
    Returns:
        pd.Series: Números de documento limpios
    """
    return (
        col.astype(str)
        .str.replace(r"\.0$", "", regex=True)  # Eliminar .0
        .str.replace(r"[^\d]", "", regex=True)  # Dejar solo dígitos
        .str.lstrip("0")  # Eliminar ceros a la izquierda
        .str.strip()
    )


def _parsear_linea(linea):
    """
    Parse a single line from transfer PDF.
    
    Args:
        linea (str): Line from transfer PDF
        
    Returns:
        dict: Parsed transfer data or None if parsing fails
    """
    partes = linea.split()
    
    if len(partes) < 4:
        return None
    
    cuenta = partes[0]
    tipo_cuenta = partes[1]
    documento = partes[2]
    
    # Find monetary value (formatted as 1,234,567.89)
    valor_idx = None
    for i, p in enumerate(partes):
        if re.match(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?", p):
            valor_idx = i
            break
    
    if valor_idx is None:
        return None
    
    valor = partes[valor_idx]
    
    # Extract name (between document and value)
    nombre = " ".join(partes[3:valor_idx]) if valor_idx > 3 else ""
    
    # Find date (DD-MM-YYYY format)
    fecha = None
    for p in reversed(partes):
        if re.match(r"\d{2}-\d{2}-\d{4}", p):
            fecha = p
            break
    
    return {
        "Cuenta": cuenta,
        "Tipo": tipo_cuenta,
        "Documento": documento,
        "Nombre": nombre,
        "Valor": valor,
        "Fecha": fecha
    }


def _normalizar_linea_ocr(linea):
    """Normaliza una línea para un matching tolerante a OCR.

    Pliega acentos a ASCII, reemplaza espacios duros (NBSP) y colapsa
    cualquier secuencia de espacios en uno solo. No altera los dígitos, solo
    homogeneiza el texto para que las búsquedas no fallen por tildes o por
    espaciado irregular del PDF/OCR.
    """
    texto = unicodedata.normalize("NFKD", str(linea or ""))
    texto = texto.encode("ascii", "ignore").decode()
    texto = texto.replace(" ", " ")
    return re.sub(r"\s+", " ", texto).strip()


def formatear_valores(cadena):
    """
    Formatea valores numéricos al formato de moneda colombiana (COP).
    
    Convierte representaciones de texto de números en cadenas monetarias formateadas
    con símbolo $, separador de miles (coma) y separador decimal (punto).
    Los múltiples valores se separan con el carácter barra vertical (|).
    
    Args:
        cadena (str or None): Cadena de entrada con números.
            Puede contener corchetes, varios números o None.
            Formato: "[1000000 2000000]" o "1000000 2000000"
    
    Returns:
        str: Cadena formateada en COP.
            Ejemplo: "$1.000.000,00 | $2.000.000,00"
            Devuelve una cadena vacía si la entrada es None o está vacía.
    
    Proceso:
        1. Valida la entrada (devuelve "" si está vacía)
        2. Elimina corchetes y espacios extra
        3. Divide múltiples valores por espacios
        4. Convierte cada valor a float y lo formatea como moneda
        5. Une múltiples valores con el separador " | "
    
    Ejemplo:
        >>> app.formatear_valores("[1000000 2500000]")
        "$1.000.000,00 | $2.500.000,00"
        
        >>> app.formatear_valores(None)
        ""
    """
    # Paso 1: validar la entrada - devolver cadena vacía si está vacía
    if not cadena:
        return ""
    
    
    # Paso 2: limpiar la cadena de entrada - quitar corchetes y espacios
    cadena = str(cadena).replace("[", "").replace("]", "").strip()
    
    # Paso 3: dividir en números individuales
    numeros = cadena.split()
    
    # Paso 4: convertir cada número a moneda formateada
    resultado = []
    for num in numeros:
        num = num.replace(",", "").strip()  # Eliminar comas existentes

        try:
            # Convertir la cadena a float
            valor = float(num)
            
            # Formatear como moneda: ${value:,.2f}
            # Ejemplo: $1000000.00 -> $1,000,000.00
            formato = "${:,.2f}".format(valor)
            
            # Convertir al formato colombiano: reemplazar comas temporalmente,
            # luego intercambiar coma y punto para obtener el formato COP
            # $1,000,000.00 → $1.000.000,00 (COP format)
            formato = formato.replace(",", "X")      # $1X000X000.00
            formato = formato.replace(".", ",")      # $1X000X000,00
            formato = formato.replace("X", ".")      # $1.000.000,00
            
            # Agregar el valor formateado al resultado
            resultado.append(formato)
            
        except (ValueError, TypeError):
            # Omitir valores que no se puedan convertir a float
            continue
    
    # Paso 5: unir múltiples valores con separador de barra vertical
    # Example: ["$1.000.000,00", "$2.000.000,00"] → "$1.000.000,00 | $2.000.000,00"
    return " | ".join(resultado)
