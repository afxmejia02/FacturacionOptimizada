"""Pruebas de extracción de equipos (``_extraer_valor_etiqueta``).

Se ejecutan sin dependencias extra:  ``python -m unittest test_equipos``

Regresión del bug: la palabra "EQUIPOS" dentro de un texto largo (la descripción
de la orden de servicio) hacía que se tomara la celda equivocada como valor del
equipo (devolvía "EQUIPO:" en vez del nombre real).
"""
import importlib.util
import os
import unittest

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
