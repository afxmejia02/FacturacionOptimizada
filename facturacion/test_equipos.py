"""Pruebas de extracción de equipos (``extraer_valor_etiqueta``).

Se ejecutan sin dependencias extra:  ``python -m unittest test_equipos``

Regresión del bug: la palabra "EQUIPOS" dentro de un texto largo (la descripción
de la orden de servicio) hacía que se tomara la celda equivocada como valor del
equipo (devolvía "EQUIPO:" en vez del nombre real).
"""
import datetime as dt
import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import facturacion as fact
from facturacion.pdf import _ETIQUETAS_EQUIPO, _ETIQUETAS_SERVICIO


ETIQUETAS = ("equipo", "tipo de equipo", "tipo equipo")


class TestExtraerValorEtiqueta(unittest.TestCase):
    def test_no_confunde_equipos_en_descripcion(self):
        # La celda de descripción contiene "EQUIPOS"; la etiqueta real es "EQUIPO:".
        row = [
            "", "EMPRESA:", None, "Consorcio Tabarca", None,
            "ORDEN DE SERVICIO No 058",
            "ALISTAMIENTO, EJECUCION ... MONTAJE DE EQUIPOS EN LA REFINERIA ...",
            None, "EQUIPO:", "CAMIÓN-GRÚA DE 10 TON (10 H)", "",
        ]
        valor = fact.extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "CAMIÓN-GRÚA DE 10 TON (10 H)")

    def test_etiqueta_y_valor_en_misma_celda(self):
        row = ["", "EQUIPO: CAMIÓN-GRÚA DE 10 TON (10 H)", ""]
        valor = fact.extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "CAMIÓN-GRÚA DE 10 TON (10 H)")

    def test_variante_tipo_de_equipo(self):
        row = ["TIPO DE EQUIPO:", "MOTOSOLDADOR HASTA 400 AMP (24 H)"]
        valor = fact.extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "MOTOSOLDADOR HASTA 400 AMP (24 H)")

    def test_conserva_texto_tras_parentesis_si_es_continuacion(self):
        # "para bridas..." es continuación legítima del nombre: NO debe recortarse.
        nombre = "Torno portátil orbital 10H (Diurno / Nocturno) para bridas >4 NPS <=\n48 NPS"
        limpio = fact.limpiar_nombre_equipo(nombre)
        self.assertIn("para bridas", limpio)
        self.assertIn("48 NPS", limpio)

    def test_recorta_fragmento_inicial_duplicado(self):
        # "Motoso" es el inicio duplicado del propio nombre (artefacto): se recorta.
        limpio = fact.limpiar_nombre_equipo("MOTOSOLDADOR HASTA 400 AMP (24 H) Motoso")
        self.assertEqual(limpio, "MOTOSOLDADOR HASTA 400 AMP (24 H)")

    def test_sin_etiqueta_devuelve_none(self):
        row = ["EMPRESA:", "Consorcio Tabarca", "ORDEN DE SERVICIO No 058"]
        self.assertIsNone(fact.extraer_valor_etiqueta([[row]], ETIQUETAS))


class TestExtraccionUnificada(unittest.TestCase):
    """Equipos y servicios comparten estructura; un solo extractor sirve a ambos."""

    def test_reconoce_etiqueta_servicio(self):
        row = ["", "SERVICIO:", "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)", ""]
        etiquetas = _ETIQUETAS_EQUIPO + _ETIQUETAS_SERVICIO
        valor = fact.extraer_valor_etiqueta([[row]], etiquetas)
        self.assertEqual(valor, "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")

    def test_extraer_registros_por_etiqueta(self):
        # Etiqueta SERVICIO: + tabla de detalle con FECHA y CANTIDAD por fila.
        tabla = [
            ["", "SERVICIO:", "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)"],
            ["ITEM", "FECHA", "IDENTIFICACION", "CANTIDAD", "UBICACIÓN"],
            ["1", "29 de mayo de 2026", "PTSC-001", "3", "GRB"],
        ]
        etiquetas = _ETIQUETAS_EQUIPO + _ETIQUETAS_SERVICIO
        registros = fact.extraer_registros_etiqueta([tabla], etiquetas)
        self.assertEqual(len(registros), 1)
        self.assertEqual(registros[0]["TIPO DE EQUIPO"], "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")
        self.assertEqual(registros[0]["CANTIDAD"], 3.0)


class TestHistogramaPorSeccion(unittest.TestCase):
    """Lectura del histograma en largo, filtrada por sección y conservando ceros."""

    def _hist_xlsx(self, fecha):
        # Dos secciones: 5.5 (equipos) y 5.6 (servicios). Incluye un ítem en 0.
        matriz = [
            ["COD. TAR.", "DESCRIPCION TARIFA", "UNIDAD", "VLR. UND.", "CANTIDAD", "VLR. TOTAL", fecha],
            ["5.5", "ELEMENTOS, HERRAMIENTAS Y EQUIPOS TRANSVERSALES", None, None, None, None, None],
            ["5.5.1.1", "CAMIÓN-GRÚA DE 10 TON (10 H)", "DÍA", 1, 1, 1, 2],
            ["5.6", "OBRAS O SERVICIOS TÍPICOS", None, None, None, None, None],
            ["5.6.2.7.1", "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)", "UN", 1, 1, 1, 5],
            ["5.6.2.7.2", "Radiografía digital computarizada", "UN", 1, 1, 1, 0],
        ]
        path = os.path.join(tempfile.mkdtemp(), "hist.xlsx")
        pd.DataFrame(matriz).to_excel(path, header=False, index=False)
        return path

    def test_filtra_por_seccion_y_conserva_ceros_y_valor_tal_cual(self):
        fecha = dt.datetime(2026, 5, 25)
        path = self._hist_xlsx(fecha)
        # Solo sección 5.6 (servicios): no debe traer el equipo de 5.5.
        largo = fact.leer_histograma_largo(path, ["5.6"])
        claves = set(largo["CLAVE"])
        self.assertIn(fact.clave_equipo("GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)"), claves)
        self.assertIn(fact.clave_equipo("Radiografía digital computarizada"), claves)  # cero conservado
        self.assertNotIn(fact.clave_equipo("CAMIÓN-GRÚA DE 10 TON (10 H)"), claves)
        # El valor se toma tal cual (sin conversión de unidades).
        gamma = largo[largo["CLAVE"] == fact.clave_equipo("GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")]
        self.assertEqual(gamma["VALOR"].iloc[0], 5)


class TestParsearCantidad(unittest.TestCase):
    """La cantidad de la planilla admite miles/decimales en cualquier convención."""

    def test_coma_decimal_sin_separador_de_miles(self):
        # Regresión: "1452,6" daba 145 (se tomaban solo 3 dígitos).
        self.assertEqual(fact.parsear_cantidad("1452,6"), 1452.6)
        self.assertEqual(fact.parsear_cantidad("1452"), 1452.0)

    def test_ambas_convenciones(self):
        self.assertEqual(fact.parsear_cantidad("1.452,6"), 1452.6)   # colombiana
        self.assertEqual(fact.parsear_cantidad("1,452.6"), 1452.6)   # anglosajona
        self.assertEqual(fact.parsear_cantidad("1.452.678"), 1452678.0)
        self.assertEqual(fact.parsear_cantidad("0,33"), 0.33)
        self.assertEqual(fact.parsear_cantidad("3"), 3.0)

    def test_vacios(self):
        self.assertIsNone(fact.parsear_cantidad(None))
        self.assertIsNone(fact.parsear_cantidad("---"))


class TestParsearCantidad(unittest.TestCase):
    """Cantidades de planilla: miles/decimales en distintas convenciones."""

    def test_convencion_colombiana(self):
        # Punto = miles, coma = decimal (misma convención en PDF y Excel).
        self.assertEqual(fact.parsear_cantidad("3.139 m³"), 3139.0)
        self.assertEqual(fact.parsear_cantidad("3.139,00"), 3139.0)
        self.assertEqual(fact.parsear_cantidad("1.452,6"), 1452.6)
        self.assertEqual(fact.parsear_cantidad("153,67 Kg"), 153.67)
        self.assertEqual(fact.parsear_cantidad("7,7 m³"), 7.7)
        self.assertEqual(fact.parsear_cantidad("1.452.678"), 1452678.0)
        self.assertEqual(fact.parsear_cantidad("3"), 3.0)
        self.assertIsNone(fact.parsear_cantidad("---"))


class TestObservacionPerfil(unittest.TestCase):
    """Parseo de la columna Observaciones: recategorización + 'E y F' + 'NO FACTURABLE'."""

    def test_marcadores_individuales(self):
        self.assertEqual(fact.parsear_observacion_perfil("E Y F"), (None, True, False, False))
        self.assertEqual(fact.parsear_observacion_perfil("NO FACTURABLE"), (None, False, True, False))
        # El patrón real trae "SE FACTURA" en medio y un salto de línea.
        self.assertEqual(
            fact.parsear_observacion_perfil("RECATEGORIZADO SE\nFACTURA COMO B4"),
            ("B4", False, False, False),
        )

    def test_24h_en_observaciones(self):
        for valor in ("24", "24H", "24HRS", "24 HORAS", "JORNADA 24 HORAS"):
            rec, ef, nf, es24 = fact.parsear_observacion_perfil(valor)
            self.assertTrue(es24, msg=repr(valor))
        # Sin "24" -> es_24h False (no debe confundirse con otros textos).
        self.assertFalse(fact.parsear_observacion_perfil("RECATEGORIZADO SE FACTURA COMO B4")[3])

    def test_coexisten_en_cualquier_orden(self):
        # Recategorización + "E y F" juntos: no importa el orden.
        self.assertEqual(
            fact.parsear_observacion_perfil("RECATEGORIZADO SE FACTURA COMO E11 E Y F"),
            ("E11", True, False, False),
        )
        self.assertEqual(
            fact.parsear_observacion_perfil("E Y F RECATEGORIZADO SE FACTURA COMO D7"),
            ("D7", True, False, False),
        )

    def test_caso_real_recat_mas_ef_con_guion(self):
        # Caso real ("13. Registro PLanillas" p.21): recat + "- E Y F" con salto de
        # línea. Debe dar el nivel recategorizado y es_ef=True (cuenta 1, no 1/3).
        self.assertEqual(
            fact.parsear_observacion_perfil("RECATEGORIZADO\nSE FACTURA COMO B4 - E Y F"),
            ("B4", True, False, False),
        )

    def test_recat_con_ruido_antes_del_nivel(self):
        # Robustez: "COMO NIVEL/PERFIL <nivel>" (palabra extra antes del nivel).
        self.assertEqual(fact.parsear_observacion_perfil("RECATEGORIZADO COMO NIVEL B4")[0], "B4")
        self.assertEqual(fact.parsear_observacion_perfil("RECATEGORIZADO COMO PERFIL C6 - E Y F")[0], "C6")

    def test_vacios_y_no_reconocidos(self):
        for valor in (None, "", "   ", "TRASLADO"):
            self.assertEqual(
                fact.parsear_observacion_perfil(valor), (None, False, False, False), msg=repr(valor)
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
