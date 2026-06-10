"""Pruebas de la extracción y conciliación de transferencias (formato ITALCO).

Se ejecutan sin dependencias extra:  ``python -m unittest test_transferencias``

Cubren la corrección del bug "Transferencia no encontrada":

- ``_match_linea_transferencia`` ahora tolera la ausencia de la columna de
  fecha, distintos formatos monetarios y ruido de OCR;
- ``_reconcile_data`` cruza por documento o cuenta y solo marca "no encontrada"
  cuando realmente no hay coincidencia.
"""
import importlib.util
import os
import unittest

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def _load(name, filename):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


payroll = _load("payroll_mod", "gui_app.py")
App = payroll.PayrollReconciliationApp


def _app():
    # __new__ evita la GUI tkinter (igual que hace la web).
    return App.__new__(App)


class TestMatchLineaTransferencia(unittest.TestCase):
    def setUp(self):
        self.app = _app()

    def test_formato_abril_sin_fecha(self):
        # El layout que rompía el patrón anterior: sin columna de fecha de 8 dígitos.
        linea = "ORTEGON GOMEZ JIMMER EDUARDO 91513843 603168089 250331 PAGO NOMINA BCA 3,484,422.00"
        data = self.app._match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91513843")
        self.assertEqual(data["Cuenta"], "603168089")
        self.assertEqual(data["Valor"], 3484422.0)

    def test_formato_con_fecha_de_8_digitos(self):
        # Layout antiguo (con fecha): debe seguir funcionando.
        linea = "PEREZ LOPEZ JUAN 91234567 057014608031 20250115 12345 PAGO NOMINA BCA 1.234.567,00"
        data = self.app._match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91234567")
        self.assertEqual(data["Cuenta"], "057014608031")
        self.assertEqual(data["Valor"], 1234567.0)

    def test_tolera_ruido_ocr_y_acentos(self):
        linea = "HERNÁNDEZ  VÁSQUEZ  ANTONIO   91433543   0570146080313076  250331   PAGO  NOMINA  BCA   5,156,483.00"
        data = self.app._match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91433543")
        self.assertEqual(data["Valor"], 5156483.0)

    def test_etiqueta_destino_flexible(self):
        # "PAGO DE NOMINA" (sin BCA) también debe reconocerse.
        linea = "GARCIA PEREZ ANA 80092626 49656405430 PAGO DE NOMINA 403,544.00"
        data = self.app._match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "80092626")
        self.assertEqual(data["Valor"], 403544.0)

    def test_lineas_que_no_son_pago_devuelven_none(self):
        for linea in (
            "BANCO DE OCCIDENTE",
            "Beneficiario Nit Beneficiario No. Producto Destino Numero Factura Vr. Pago",
            "Empresa: UNION TEMPORAL ITALCO",
            "",
            None,
        ):
            self.assertIsNone(self.app._match_linea_transferencia(linea), msg=repr(linea))


class TestReconcileData(unittest.TestCase):
    def setUp(self):
        self.app = _app()

    def _despr(self, filas):
        return pd.DataFrame(filas)

    def test_cruce_por_documento_no_es_no_encontrada(self):
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3484422, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = self.app._reconcile_data(despr, trans, None)
        estado = df_t.iloc[0]["Estado"]
        self.assertNotEqual(estado, "Transferencia no encontrada")
        self.assertEqual(estado, "OK")  # suma de netos == suma de transferencias

    def test_suma_distinta_es_valor_no_coincide(self):
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3457617, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = self.app._reconcile_data(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "Valor no coincide")

    def test_sin_coincidencia_es_no_encontrada(self):
        despr = self._despr([
            {"Identificacion": "99999999", "Neto": 1000, "Devengado": None, "Cuenta": "111"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = self.app._reconcile_data(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "Transferencia no encontrada")

    def test_cruce_por_cuenta_con_ceros_de_relleno(self):
        # El documento difiere pero la cuenta coincide salvo ceros a la izquierda.
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3484422, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "0000000", "Cuenta": "00603168089", "Valor": 3484422},
        ])
        df_t, _ = self.app._reconcile_data(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "OK")


if __name__ == "__main__":
    unittest.main(verbosity=2)
