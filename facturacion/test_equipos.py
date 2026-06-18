"""Pruebas de extracción de equipos (``_extraer_valor_etiqueta``).

Se ejecutan sin dependencias extra:  ``python -m unittest test_equipos``

Regresión del bug: la palabra "EQUIPOS" dentro de un texto largo (la descripción
de la orden de servicio) hacía que se tomara la celda equivocada como valor del
equipo (devolvía "EQUIPO:" en vez del nombre real).
"""
import datetime as dt
import importlib.util
import os
import tempfile
import unittest

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


validator = _load("validator_mod", "gui_validation_app.py")


def _app():
    app = validator.ServicesValidationApp.__new__(validator.ServicesValidationApp)
    app.debug_mode = False
    return app


ETIQUETAS = ("equipo", "tipo de equipo", "tipo equipo")


class TestExtraerValorEtiqueta(unittest.TestCase):
    def setUp(self):
        self.app = _app()

    def test_no_confunde_equipos_en_descripcion(self):
        # La celda de descripción contiene "EQUIPOS"; la etiqueta real es "EQUIPO:".
        row = [
            "", "EMPRESA:", None, "Consorcio Tabarca", None,
            "ORDEN DE SERVICIO No 058",
            "ALISTAMIENTO, EJECUCION ... MONTAJE DE EQUIPOS EN LA REFINERIA ...",
            None, "EQUIPO:", "CAMIÓN-GRÚA DE 10 TON (10 H)", "",
        ]
        valor = self.app._extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "CAMIÓN-GRÚA DE 10 TON (10 H)")

    def test_etiqueta_y_valor_en_misma_celda(self):
        row = ["", "EQUIPO: CAMIÓN-GRÚA DE 10 TON (10 H)", ""]
        valor = self.app._extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "CAMIÓN-GRÚA DE 10 TON (10 H)")

    def test_variante_tipo_de_equipo(self):
        row = ["TIPO DE EQUIPO:", "MOTOSOLDADOR HASTA 400 AMP (24 H)"]
        valor = self.app._extraer_valor_etiqueta([[row]], ETIQUETAS)
        self.assertEqual(valor, "MOTOSOLDADOR HASTA 400 AMP (24 H)")

    def test_sin_etiqueta_devuelve_none(self):
        row = ["EMPRESA:", "Consorcio Tabarca", "ORDEN DE SERVICIO No 058"]
        self.assertIsNone(self.app._extraer_valor_etiqueta([[row]], ETIQUETAS))


class TestExtraccionUnificada(unittest.TestCase):
    """Equipos y servicios comparten estructura; un solo extractor sirve a ambos."""

    def setUp(self):
        self.app = _app()

    def test_reconoce_etiqueta_servicio(self):
        row = ["", "SERVICIO:", "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)", ""]
        etiquetas = self.app._ETIQUETAS_EQUIPO + self.app._ETIQUETAS_SERVICIO
        valor = self.app._extraer_valor_etiqueta([[row]], etiquetas)
        self.assertEqual(valor, "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")

    def test_extraer_registros_por_etiqueta(self):
        # Etiqueta SERVICIO: + tabla de detalle con FECHA y CANTIDAD por fila.
        tabla = [
            ["", "SERVICIO:", "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)"],
            ["ITEM", "FECHA", "IDENTIFICACION", "CANTIDAD", "UBICACIÓN"],
            ["1", "29 de mayo de 2026", "PTSC-001", "3", "GRB"],
        ]
        etiquetas = self.app._ETIQUETAS_EQUIPO + self.app._ETIQUETAS_SERVICIO
        registros = self.app._extraer_registros_etiqueta([tabla], etiquetas)
        self.assertEqual(len(registros), 1)
        self.assertEqual(registros[0]["TIPO DE EQUIPO"], "GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")
        self.assertEqual(registros[0]["CANTIDAD"], 3.0)


class TestConversionUnidadMes(unittest.TestCase):
    """Tarifas en 'MES' se pasan a unidad diaria (× 30) y se aproximan al entero."""

    def setUp(self):
        self.app = _app()

    def _hist_xlsx(self, fecha):
        # Encabezado real con columna UNIDAD + una columna de fecha (datetime).
        matriz = [
            ["COD. TAR.", "DESCRIPCION TARIFA", "UNIDAD", "VLR. UND.", "CANTIDAD", "VLR. TOTAL", fecha],
            ["1.1", "CAMPEROS O CAMIONETAS 4X2 (10 HORAS)", "MES", 1, 1, 1, 0.099],
            ["1.2", "Nivel E11", "DÍA", 1, 1, 1, 11],
        ]
        path = os.path.join(tempfile.mkdtemp(), "hist.xlsx")
        pd.DataFrame(matriz).to_excel(path, header=False, index=False)
        return path

    def test_mes_a_diaria_y_dia_sin_cambios(self):
        fecha = dt.datetime(2026, 5, 25)
        path = self._hist_xlsx(fecha)
        d = self.app._extraer_conteo_excel(path, fecha)
        # 0.099 × 30 = 2.97 -> 3
        self.assertEqual(d[self.app._clave_equipo("CAMPEROS O CAMIONETAS 4X2 (10 HORAS)")], 3.0)
        # Las tarifas en DÍA no se alteran.
        self.assertEqual(d[self.app._clave_equipo("Nivel E11")], 11.0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
