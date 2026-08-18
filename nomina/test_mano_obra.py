"""Pruebas de la comparación de mano de obra (Informe de Costo vs ODS).

Se ejecutan sin dependencias extra:  ``python -m unittest test_mano_obra``

Cubren, además de los helpers, las dos correcciones recientes:

1. las fechas de actividades de la ODS se comparan contra ``Fecha Inicio`` /
   ``Fecha Vencimiento`` del Informe (no contra ``Fecha de Ingreso`` / retiro);
2. la comparación de salario (COP) y la ausencia de la columna de estado.
"""
import datetime as dt
import importlib.util
import io
import os
import unittest

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


mo = _load("mano_obra_mod", "mano_obra.py")


def _informe_xlsx(filas):
    """Crea un Informe en memoria con el encabezado real en la fila 10 (índice 9).

    ``filas`` es una lista de dicts con las columnas que usamos por posición.
    """
    n_cols = 18  # se usan posiciones hasta la 17 (Salario Diario Contratado)
    filler = [[None] * n_cols for _ in range(9)]          # filas 0-8 (títulos)
    # Encabezado real (fila 9). Las columnas se localizan por NOMBRE, no posición.
    header = [f"c{i}" for i in range(n_cols)]
    header[0] = "Valor pagado en nomina"
    header[2] = "Identificación"
    header[4] = "Nombres"
    header[5] = "Apellidos"
    header[6] = "Cargo"
    header[7] = "Fecha Inicio"
    header[8] = "Fecha Vencimiento"
    header[10] = "Nombre Centro Costo"
    header[16] = "Días Trabajados"
    header[17] = "Salario Diario Contratado"
    datos = []
    for f in filas:
        fila = [None] * n_cols
        fila[0] = f.get("tipo", "Recobro")
        fila[2] = f["doc"]
        fila[4] = f.get("nombres", "Juan")
        fila[5] = f.get("apellidos", "Perez")
        fila[6] = f.get("cargo", "Almacenista")
        fila[7] = f.get("fecha_inicio")
        fila[8] = f.get("fecha_vencimiento")
        fila[10] = f.get("centro", "Ecop Os050 - X")
        fila[16] = f.get("dias", 30)
        fila[17] = f.get("salario")
        datos.append(fila)

    matriz = filler + [header] + datos
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(matriz).to_excel(writer, sheet_name="Informe", header=False, index=False)
    buffer.seek(0)
    return buffer


def _ods_xlsx(filas):
    col_ini = "Fecha_de_inicio_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio"
    col_fin = "Fecha_fin_de_actividades_del_trabajador_para_el_contrato_comercial_u_orden_de_servicio"
    registros = []
    for f in filas:
        registros.append({
            "NumeroDocumento": f["doc"],
            "No_de_orden_de_servicio_conocido_por_el_contratista": f.get("os", 50),
            "Nombres": f.get("nombres", "Juan"),
            "Apellidos": f.get("apellidos", "Perez"),
            "CargoContratoLaboral": f.get("cargo", "Almacenista"),
            col_ini: f.get("fecha_inicio"),
            col_fin: f.get("fecha_fin"),
            "DiasTrabajadosEnMes": f.get("dias", 30),
            "SalarioDiarioPesos": f.get("salario"),
        })
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(registros).to_excel(writer, index=False)
    buffer.seek(0)
    return buffer


def _informe_italco_xlsx(filas, columna_os=False):
    """Crea un Informe ITALCO (progresión) en memoria.

    El encabezado real no es la primera fila (hay filas de título arriba) y la
    primera columna no tiene nombre: marca ACTUAL / ANTERIOR / DIFERENCIA. Por
    cada persona se emiten las tres filas; solo la de DIFERENCIA debe usarse.

    ``columna_os`` reproduce el otro layout visto en producción: en vez de
    ``Perfil Contable`` con el texto ``BCA OS 37 CONVENCIONAL``, el exporte trae
    una columna ``OS`` con el número suelto (y numérico, o sea float).
    """
    titulos = [
        [None, "Union temporal Italco"],
        [None, "PROYECTO UT Barranca"],
        [None, "INFORMACION DE NOMINA"],
    ]
    header = [
        None, "Mes", "Documento", "Nombre Completo", "Tipo de Contrato",
        "Fecha de Inicio", "Fecha retiro", "OS" if columna_os else "Perfil Contable",
        "Cargo", "Sueldo Base",
    ]
    datos = []
    for f in filas:
        # El Sueldo Base solo importa en la fila ACTUAL (de ahí sale el salario
        # diario); en ANTERIOR/DIFERENCIA se pone un delta cualquiera.
        sueldo = {"ACTUAL": f.get("sueldo_base_actual"), "ANTERIOR": 0, "DIFERENCIA": 123}
        os_valor = f.get("os_num", 37.0) if columna_os else f.get("perfil", "BCA OS 37 CONVENCIONAL")
        for marca in ("ACTUAL", "ANTERIOR", "DIFERENCIA"):
            datos.append([
                marca, "2025 7", f["doc"], f.get("nombre", "JUAN PEREZ"),
                "ECP R", f.get("fecha_inicio"), f.get("fecha_retiro"),
                os_valor, f.get("cargo", "OBRERO A"),
                sueldo[marca],
            ])
    matriz = titulos + [header] + datos
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(matriz).to_excel(
            writer, sheet_name="PROGRESION JULIO 2025", header=False, index=False
        )
    buffer.seek(0)
    return buffer


class TestNormalizarMoneda(unittest.TestCase):
    def test_numericos_y_strings_equivalentes(self):
        for valor in (120000, 120000.0, "120000", "$120.000", "120,000", "  $ 120.000 "):
            self.assertEqual(mo.normalizar_moneda(valor), 120000.0, msg=repr(valor))

    def test_decimales_ambas_convenciones(self):
        self.assertEqual(mo.normalizar_moneda("1.234.567,89"), 1234567.89)
        self.assertEqual(mo.normalizar_moneda("1,234,567.89"), 1234567.89)
        self.assertEqual(mo.normalizar_moneda("120.5"), 120.5)

    def test_vacios_e_invalidos(self):
        for valor in (None, float("nan"), "", "   ", "abc", True, False):
            self.assertIsNone(mo.normalizar_moneda(valor), msg=repr(valor))


class TestSoloDigitos(unittest.TestCase):
    """Regresión: documentos que Excel entrega como número no deben ganar un 0."""

    def test_float_entero_no_arrastra_el_decimal(self):
        # str(1096198448.0) == "1096198448.0"; quitar el punto pegaría el 0 final.
        self.assertEqual(mo.solo_digitos(1096198448.0), "1096198448")
        self.assertEqual(mo.solo_digitos(91517631.0), "91517631")

    def test_texto_con_separadores_de_miles(self):
        self.assertEqual(mo.solo_digitos("1.096.216.566"), "1096216566")
        self.assertEqual(mo.solo_digitos("91.499.442"), "91499442")

    def test_float_y_texto_del_mismo_documento_coinciden(self):
        # Informe (número) y ODS (texto) deben reducirse al mismo documento.
        self.assertEqual(mo.solo_digitos(1096216566.0), mo.solo_digitos("1.096.216.566"))

    def test_enteros_y_vacios(self):
        self.assertEqual(mo.solo_digitos(91499442), "91499442")
        self.assertEqual(mo.solo_digitos(""), "")


class TestFormatearCop(unittest.TestCase):
    def test_entero_sin_decimales(self):
        self.assertEqual(mo.formatear_cop(120000), "$120.000")
        self.assertEqual(mo.formatear_cop(1234567), "$1.234.567")

    def test_con_decimales(self):
        self.assertEqual(mo.formatear_cop(120000.5), "$120.000,50")

    def test_vacio(self):
        self.assertEqual(mo.formatear_cop(None), "")
        self.assertEqual(mo.formatear_cop(""), "")


class TestComparable(unittest.TestCase):
    def test_moneda_iguales_distinto_formato(self):
        self.assertEqual(mo._comparable("$120.000", "moneda"), mo._comparable(120000, "moneda"))

    def test_moneda_distintas(self):
        self.assertNotEqual(mo._comparable(120000, "moneda"), mo._comparable(125000, "moneda"))

    def test_fecha_timestamp_vs_iso_y_datetime(self):
        ts = pd.Timestamp("2025-06-01")
        self.assertEqual(mo._comparable(ts, "fecha"), mo._comparable("2025-06-01", "fecha"))
        self.assertEqual(mo._comparable(ts, "fecha"), mo._comparable(dt.datetime(2025, 6, 1, 9, 30), "fecha"))

    def test_fecha_nulos_se_consideran_iguales(self):
        self.assertEqual(mo._comparable(None, "fecha"), mo._comparable(float("nan"), "fecha"))


class TestMapeoColumnas(unittest.TestCase):
    """Regresión del bug: las fechas de actividades ODS ↔ Fecha Inicio/Vencimiento."""

    def _entrada(self, etiqueta):
        return next(e for e in mo.MAPEO_COLUMNAS if e[0] == etiqueta)

    def test_fecha_inicio_contra_inicio_actividades(self):
        etiqueta, col_inf, col_ods, tipo = self._entrada("Fecha Inicio")
        self.assertEqual(col_inf, "Fecha Inicio")
        self.assertIn("inicio_de_actividades", col_ods)
        self.assertEqual(tipo, "fecha")

    def test_fecha_vencimiento_contra_fin_actividades(self):
        etiqueta, col_inf, col_ods, tipo = self._entrada("Fecha Vencimiento")
        self.assertEqual(col_inf, "Fecha Vencimiento")
        self.assertIn("fin_de_actividades", col_ods)

    def test_no_se_compara_contra_ingreso_ni_retiro(self):
        cols_informe = {e[1] for e in mo.MAPEO_COLUMNAS}
        self.assertNotIn("Fecha de Ingreso", cols_informe)
        self.assertNotIn("Fecha de retiro", cols_informe)

    def test_existe_comparacion_salario_moneda(self):
        etiqueta, col_inf, col_ods, tipo = self._entrada("Salario")
        self.assertEqual(col_inf, "Salario Diario Contratado")
        self.assertEqual(col_ods, "SalarioDiarioPesos")
        self.assertEqual(tipo, "moneda")


class TestCompararManoObra(unittest.TestCase):
    def test_end_to_end_fechas_salario_y_sin_estado(self):
        # Mismo doc en ambos. Fecha de inicio del contrato COINCIDE; el
        # vencimiento DIFIERE; el salario COINCIDE.
        informe = _informe_xlsx([{
            "doc": "91499442",
            "fecha_inicio": dt.datetime(2024, 1, 15),
            "fecha_vencimiento": dt.datetime(2025, 12, 31),
            "salario": 120000,
        }])
        ods = _ods_xlsx([{
            "doc": "91499442",
            "fecha_inicio": dt.datetime(2024, 1, 15),   # == Fecha Inicio del Informe
            "fecha_fin": dt.datetime(2025, 11, 30),      # != Fecha Vencimiento
            "salario": 120000,
        }])
        df = mo.comparar_mano_obra(informe, ods)

        # No hay columna de estado/observaciones.
        self.assertNotIn(mo.COL_ESTADO if hasattr(mo, "COL_ESTADO") else "Estado revisión", df.columns)
        self.assertNotIn("Estado revisión", df.columns)
        self.assertNotIn("Observaciones", df.columns)

        fila = df.iloc[0]
        # Fecha de inicio coincide -> lista de un elemento.
        self.assertEqual(fila["Fecha Inicio"], ["2024-01-15"])
        # Fecha de vencimiento difiere -> dos elementos.
        self.assertEqual(fila["Fecha Vencimiento"], ["2025-12-31", "2025-11-30"])
        # Salario coincide -> un elemento, formateado en COP.
        self.assertEqual(fila["Salario"], ["$120.000"])

    def test_salario_distinto_muestra_ambos_en_cop(self):
        informe = _informe_xlsx([{
            "doc": "91499442",
            "fecha_inicio": dt.datetime(2024, 1, 15),
            "fecha_vencimiento": dt.datetime(2025, 12, 31),
            "salario": 120000,
        }])
        ods = _ods_xlsx([{
            "doc": "91499442",
            "fecha_inicio": dt.datetime(2024, 1, 15),
            "fecha_fin": dt.datetime(2025, 12, 31),
            "salario": 125000,
        }])
        df = mo.comparar_mano_obra(informe, ods)
        self.assertEqual(df.iloc[0]["Salario"], ["$120.000", "$125.000"])

    def test_acepta_varios_informes_y_ods(self):
        # Cada persona viene en un archivo distinto por lado: el cruce debe
        # encontrar ambas al concatenar los Excel.
        informe_a = _informe_xlsx([{"doc": "111", "salario": 100000}])
        informe_b = _informe_xlsx([{"doc": "222", "salario": 200000}])
        ods_a = _ods_xlsx([{"doc": "111", "salario": 100000}])
        ods_b = _ods_xlsx([{"doc": "222", "salario": 200000}])

        df = mo.comparar_mano_obra([informe_a, informe_b], [ods_a, ods_b])

        docs = set(df[mo.COL_DOCUMENTO].astype(str))
        self.assertEqual(docs, {"111", "222"})
        self.assertEqual(len(df), 2)


class TestMapeoItalco(unittest.TestCase):
    """El mapeo ITALCO compara nombre completo, cargo y las dos fechas de actividades."""

    def _entrada(self, etiqueta):
        return next(e for e in mo.MAPEO_ITALCO if e[0] == etiqueta)

    def test_fecha_inicio_contra_inicio_actividades(self):
        _, col_inf, col_ods, tipo = self._entrada("Fecha Inicio")
        self.assertEqual(col_inf, "Fecha de Inicio")
        self.assertIn("inicio_de_actividades", col_ods)
        self.assertEqual(tipo, "fecha")

    def test_fecha_retiro_contra_fin_actividades(self):
        _, col_inf, col_ods, tipo = self._entrada("Fecha Retiro")
        self.assertEqual(col_inf, "Fecha retiro")
        self.assertIn("fin_de_actividades", col_ods)

    def test_nombre_completo_contra_combinado_ods(self):
        _, col_inf, col_ods, _ = self._entrada("Nombre Completo")
        self.assertEqual(col_inf, "Nombre Completo")
        self.assertEqual(col_ods, mo.COL_NOMBRE_COMPLETO_ODS)

    def test_cargo_usa_tipo_cargo(self):
        _, _, _, tipo = self._entrada("Cargo")
        self.assertEqual(tipo, "cargo")

    def test_os_contra_orden_servicio(self):
        _, col_inf, col_ods, tipo = self._entrada("OS")
        self.assertEqual(col_inf, "OS")
        self.assertIn("orden_de_servicio", col_ods)
        self.assertEqual(tipo, "os")


class TestOsComparable(unittest.TestCase):
    def test_extrae_numero_de_ambos_formatos(self):
        # Informe "BCA OS 37 CONVENCIONAL" y ODS "0DS37" deben coincidir en 37.
        self.assertEqual(mo._os_comparable("BCA OS 37 CONVENCIONAL"), "37")
        self.assertEqual(mo._os_comparable("0DS37"), "37")
        self.assertEqual(mo._os_comparable("BCA OS 37 CONVENCIONAL"), mo._os_comparable("0DS37"))

    def test_distinta_os_difiere(self):
        self.assertNotEqual(mo._os_comparable("BCA OS 37"), mo._os_comparable("0DS44"))

    def test_vacio(self):
        self.assertEqual(mo._os_comparable(None), "")
        self.assertEqual(mo._os_comparable("SIN NUMERO"), "")

    def test_numero_suelto_como_float(self):
        # Exportes con columna "OS" numérica: str(37.0) == "37.0" y, al tomar el
        # último grupo de dígitos, devolvería "0" en vez de "37".
        self.assertEqual(mo._os_comparable(37.0), "37")
        self.assertEqual(mo._os_comparable(37), "37")
        self.assertEqual(mo._os_comparable(37.0), mo._os_comparable("0DS37"))


class TestNormCargo(unittest.TestCase):
    def test_ignora_marcador_progresion(self):
        # El "(PROGRE)" del Informe no debe contar como diferencia.
        self.assertEqual(
            mo.norm_cargo("AYUDANTE TECNICO A / TUBERIA C6 (PROGRE)"),
            mo.norm_cargo("AYUDANTE TECNICO  A / TUBERIA C6"),
        )

    def test_conserva_diferencias_reales(self):
        # E12 vs E11 sigue siendo una diferencia real.
        self.assertNotEqual(
            mo.norm_cargo("ELECTRICISTA 1A E12 (PROGRE)"),
            mo.norm_cargo("ELECTRICISTA 1A E11"),
        )


class TestCoincidenCargo(unittest.TestCase):
    """Un cargo del Informe con alternativas 'A / B' coincide si la ODS matchea una."""

    def test_ods_matchea_una_alternativa(self):
        # ODS "ANDAMIERO B" coincide con la primera alternativa de "ANDAMIERO B / D8".
        self.assertTrue(mo._coinciden("ANDAMIERO B / D8 (PROGRE)", "ANDAMIERO B", "cargo"))

    def test_ods_con_forma_completa_tambien_coincide(self):
        self.assertTrue(mo._coinciden(
            "AYUDANTE TECNICO A / TUBERIA C6 (PROGRE)",
            "AYUDANTE TECNICO  A / TUBERIA C6",
            "cargo",
        ))

    def test_sin_alternativa_comun_difiere(self):
        self.assertFalse(mo._coinciden("OBRERO A (PROGRE)", "SOLDADOR B", "cargo"))

    def test_uno_vacio_difiere(self):
        self.assertFalse(mo._coinciden("ANDAMIERO B / D8", "", "cargo"))
        self.assertTrue(mo._coinciden("", None, "cargo"))


class TestCompararManoObraItalco(unittest.TestCase):
    def test_solo_usa_filas_diferencia_y_cruza_fechas(self):
        # Documento presente en ambos. Nombre y cargo coinciden; la fecha de
        # inicio coincide; la de retiro difiere (Informe vacío vs ODS con fecha).
        informe = _informe_italco_xlsx([{
            "doc": "91.449.953",
            "nombre": "ALBEIRO RODRIGUEZ DURAN",
            "cargo": "AYUDANTE TECNICO A",
            "perfil": "BCA OS 37 CONVENCIONAL",
            "sueldo_base_actual": 4077630,  # / 30 = 135.921 (salario diario)
            "fecha_inicio": dt.datetime(2024, 8, 26),
            "fecha_retiro": None,
        }])
        ods = _ods_xlsx([{
            "doc": "91449953",
            "os": "0DS37",  # la ODS trae la OS como "0DS37"
            "nombres": "ALBEIRO",
            "apellidos": "RODRIGUEZ DURAN",
            "cargo": "AYUDANTE TECNICO A",
            "salario": 135921,  # SalarioDiarioPesos == Sueldo Base ACTUAL / 30
            "fecha_inicio": dt.datetime(2024, 8, 26),
            "fecha_fin": dt.datetime(2025, 8, 30),
        }])
        df = mo.comparar_mano_obra(informe, ods, formato="italco")

        # Una sola persona (las filas ACTUAL/ANTERIOR se descartan).
        self.assertEqual(len(df), 1)
        fila = df.iloc[0]
        self.assertEqual(
            list(df.columns),
            ["Documento", "OS", "Nombre Completo", "Cargo", "Fecha Inicio", "Fecha Retiro", "Salario"],
        )
        # OS coincide pese al formato distinto (BCA OS 37 CONVENCIONAL vs 0DS37).
        self.assertEqual(fila["OS"], ["37"])
        # Nombre y cargo coinciden -> un elemento.
        self.assertEqual(fila["Nombre Completo"], ["ALBEIRO RODRIGUEZ DURAN"])
        self.assertEqual(fila["Cargo"], ["AYUDANTE TECNICO A"])
        # Fecha inicio coincide, retiro difiere (vacío vs fecha).
        self.assertEqual(fila["Fecha Inicio"], ["2024-08-26"])
        self.assertEqual(fila["Fecha Retiro"], ["", "2025-08-30"])
        # Salario diario = Sueldo Base ACTUAL / 30, coincide con la ODS, en COP.
        self.assertEqual(fila["Salario"], ["$135.921"])

    def test_salario_diario_deriva_de_actual_no_del_delta(self):
        # El salario diario sale del Sueldo Base de la fila ACTUAL / 30, NO de la
        # columna de la fila DIFERENCIA (que es un delta). Aquí difiere de la ODS.
        informe = _informe_italco_xlsx([{
            "doc": "5",
            "sueldo_base_actual": 4598550,  # / 30 = 153.285
        }])
        ods = _ods_xlsx([{"doc": "5", "salario": 200000}])  # ODS distinta
        df = mo.comparar_mano_obra(informe, ods, formato="italco")
        # Muestra el diario derivado (153.285), no el delta (123) ni el mensual.
        self.assertEqual(df.iloc[0]["Salario"], ["$153.285", "$200.000"])

    def test_incluye_persona_del_informe_sin_ods(self):
        # ITALCO lista a todos los del Informe aunque no estén en la ODS: la
        # persona aparece con el lado ODS en blanco (celdas resaltadas).
        informe = _informe_italco_xlsx([{
            "doc": "111",
            "nombre": "PEDRO GOMEZ",
            "cargo": "OBRERO A",
            "fecha_inicio": dt.datetime(2025, 7, 1),
        }])
        ods = _ods_xlsx([{"doc": "222"}])  # otra persona: 111 no está en la ODS
        df = mo.comparar_mano_obra(informe, ods, formato="italco")

        self.assertEqual(list(df[mo.COL_DOCUMENTO]), ["111"])
        fila = df.iloc[0]
        # Cada campo con valor en el Informe queda como diferencia (ODS vacío).
        self.assertEqual(fila["Nombre Completo"], ["PEDRO GOMEZ", ""])
        self.assertEqual(fila["Cargo"], ["OBRERO A", ""])
        self.assertEqual(fila["Fecha Inicio"], ["2025-07-01", ""])

    def test_documento_numerico_cruza_contra_ods_en_texto(self):
        # Regresión: el Informe trae el documento como número (float, por algún
        # vacío en la columna) y la ODS como texto con puntos de miles. Si no se
        # normaliza el float, el documento gana un 0 final y no cruza nadie.
        informe = _informe_italco_xlsx([{
            "doc": 1096198448.0,
            "nombre": "BRIAN MADERA",
            "cargo": "RESCATISTA B3",
        }])
        ods = _ods_xlsx([{
            "doc": "1.096.198.448",
            "nombres": "BRIAN",
            "apellidos": "MADERA",
            "cargo": "RESCATISTA B3",
        }])
        df = mo.comparar_mano_obra(informe, ods, formato="italco")

        self.assertEqual(list(df[mo.COL_DOCUMENTO]), ["1096198448"])
        # Cruzó: el nombre y el cargo coinciden -> un solo elemento.
        self.assertEqual(df.iloc[0]["Nombre Completo"], ["BRIAN MADERA"])
        self.assertEqual(df.iloc[0]["Cargo"], ["RESCATISTA B3"])

    def test_informe_con_columna_os_numerica(self):
        # Otro layout: la OS viene como columna "OS" numérica (37.0) en vez de
        # dentro del "Perfil Contable". Debe cruzar contra el "0DS37" de la ODS.
        informe = _informe_italco_xlsx([{"doc": "91449953"}], columna_os=True)
        ods = _ods_xlsx([{"doc": "91449953", "os": "0DS37"}])
        df = mo.comparar_mano_obra(informe, ods, formato="italco")
        self.assertEqual(df.iloc[0]["OS"], ["37"])

    def test_tabarca_tambien_incluye_no_cruzados(self):
        # TABARCA también lista a quien está en el Informe pero no en la ODS.
        informe = _informe_xlsx([{"doc": "111", "salario": 100000}])
        ods = _ods_xlsx([{"doc": "222"}])  # 111 no está en la ODS
        df = mo.comparar_mano_obra(informe, ods)
        self.assertEqual(list(df[mo.COL_DOCUMENTO]), ["111"])
        # El salario del Informe queda como diferencia (ODS vacío).
        self.assertEqual(df.iloc[0]["Salario"], ["$100.000", ""])


if __name__ == "__main__":
    unittest.main(verbosity=2)
