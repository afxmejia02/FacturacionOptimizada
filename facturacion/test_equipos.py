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


class TestHistogramaPorSeccion(unittest.TestCase):
    """Lectura del histograma en largo, filtrada por sección y conservando ceros."""

    def setUp(self):
        self.app = _app()

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
        largo = self.app._leer_histograma_largo(path, ["5.6"])
        claves = set(largo["CLAVE"])
        self.assertIn(self.app._clave_equipo("GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)"), claves)
        self.assertIn(self.app._clave_equipo("Radiografía digital computarizada"), claves)  # cero conservado
        self.assertNotIn(self.app._clave_equipo("CAMIÓN-GRÚA DE 10 TON (10 H)"), claves)
        # El valor se toma tal cual (sin conversión de unidades).
        gamma = largo[largo["CLAVE"] == self.app._clave_equipo("GAMMAGRAFÍAS (RADIOGRAFÍA CONVENCIONAL)")]
        self.assertEqual(gamma["VALOR"].iloc[0], 5)


class TestObservacionPerfil(unittest.TestCase):
    """Parseo de la columna Observaciones: recategorización + 'E y F' + 'NO FACTURABLE'."""

    def setUp(self):
        self.app = _app()

    def test_marcadores_individuales(self):
        self.assertEqual(self.app._parsear_observacion_perfil("E Y F"), (None, True, False, False))
        self.assertEqual(self.app._parsear_observacion_perfil("NO FACTURABLE"), (None, False, True, False))
        # El patrón real trae "SE FACTURA" en medio y un salto de línea.
        self.assertEqual(
            self.app._parsear_observacion_perfil("RECATEGORIZADO SE\nFACTURA COMO B4"),
            ("B4", False, False, False),
        )

    def test_24h_en_observaciones(self):
        for valor in ("24", "24H", "24HRS", "24 HORAS", "JORNADA 24 HORAS"):
            rec, ef, nf, es24 = self.app._parsear_observacion_perfil(valor)
            self.assertTrue(es24, msg=repr(valor))
        # Sin "24" -> es_24h False (no debe confundirse con otros textos).
        self.assertFalse(self.app._parsear_observacion_perfil("RECATEGORIZADO SE FACTURA COMO B4")[3])

    def test_coexisten_en_cualquier_orden(self):
        # Recategorización + "E y F" juntos: no importa el orden.
        self.assertEqual(
            self.app._parsear_observacion_perfil("RECATEGORIZADO SE FACTURA COMO E11 E Y F"),
            ("E11", True, False, False),
        )
        self.assertEqual(
            self.app._parsear_observacion_perfil("E Y F RECATEGORIZADO SE FACTURA COMO D7"),
            ("D7", True, False, False),
        )

    def test_vacios_y_no_reconocidos(self):
        for valor in (None, "", "   ", "TRASLADO"):
            self.assertEqual(
                self.app._parsear_observacion_perfil(valor), (None, False, False, False), msg=repr(valor)
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
