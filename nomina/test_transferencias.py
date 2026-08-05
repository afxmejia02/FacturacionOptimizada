"""Pruebas de la extracción y conciliación de transferencias (formato ITALCO).

Se ejecutan sin dependencias extra:  ``python -m unittest test_transferencias``

Cubren la corrección del bug "Transferencia no encontrada":

- ``match_linea_transferencia`` ahora tolera la ausencia de la columna de
  fecha, distintos formatos monetarios y ruido de OCR;
- ``conciliar`` cruza por documento o cuenta y solo marca "no encontrada"
  cuando realmente no hay coincidencia.
"""
import os
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from nomina import conciliar, match_linea_transferencia


class TestMatchLineaTransferencia(unittest.TestCase):
    def test_formato_abril_sin_fecha(self):
        # El layout que rompía el patrón anterior: sin columna de fecha de 8 dígitos.
        linea = "ORTEGON GOMEZ JIMMER EDUARDO 91513843 603168089 250331 PAGO NOMINA BCA 3,484,422.00"
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91513843")
        self.assertEqual(data["Cuenta"], "603168089")
        self.assertEqual(data["Valor"], 3484422.0)

    def test_formato_con_fecha_de_8_digitos(self):
        # Layout antiguo (con fecha): debe seguir funcionando.
        linea = "PEREZ LOPEZ JUAN 91234567 057014608031 20250115 12345 PAGO NOMINA BCA 1.234.567,00"
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91234567")
        self.assertEqual(data["Cuenta"], "057014608031")
        self.assertEqual(data["Valor"], 1234567.0)

    def test_tolera_ruido_ocr_y_acentos(self):
        linea = "HERNÁNDEZ  VÁSQUEZ  ANTONIO   91433543   0570146080313076  250331   PAGO  NOMINA  BCA   5,156,483.00"
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91433543")
        self.assertEqual(data["Valor"], 5156483.0)

    def test_etiqueta_destino_flexible(self):
        # "PAGO DE NOMINA" (sin BCA) también debe reconocerse.
        linea = "GARCIA PEREZ ANA 80092626 49656405430 PAGO DE NOMINA 403,544.00"
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "80092626")
        self.assertEqual(data["Valor"], 403544.0)

    def test_valor_ignora_columna_ods_final(self):
        # Formato con columna "ods" al final: el valor NO es el último número.
        linea = (
            "00000000670833904 ORTEGON GOMEZ JIMMER EDUARDO 91513843 603168089 "
            "20250505 250430 PAGO NOMINA BCA 3,457,617.00 40"
        )
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "91513843")
        self.assertEqual(data["Valor"], 3457617.0)  # no 40 (la columna ods)

    def test_extrae_fecha_de_factura(self):
        # La fecha = número de factura (YYMMDD) antes de "PAGO".
        linea = "ORTEGON GOMEZ JIMMER EDUARDO 91513843 603168089 250415 PAGO NOMINA BCA 2,447,732.00"
        data = match_linea_transferencia(linea)
        self.assertEqual(data["Fecha"], pd.Timestamp("2025-04-15"))
        self.assertTrue(data["EsNomina"])

    def test_candidato_sin_etiqueta_nomina_se_extrae(self):
        # Layout "consulta de pagos a terceros": sin etiqueta NÓMINA ni fecha de
        # quincena. Debe extraerse como candidato (EsNomina=False, Fecha=None).
        linea = (
            "00000000670833904 SOTO JARABA JOSE MIGUEL 1005179167 259152502 "
            "670F0422518500DZ 20250704 2,677,442.00"
        )
        data = match_linea_transferencia(linea)
        self.assertIsNotNone(data)
        self.assertEqual(data["Documento"], "1005179167")
        self.assertEqual(data["Valor"], 2677442.0)
        self.assertFalse(data["EsNomina"])
        self.assertIsNone(data["Fecha"])

    def test_lineas_que_no_son_pago_devuelven_none(self):
        for linea in (
            "BANCO DE OCCIDENTE",
            "Beneficiario Nit Beneficiario No. Producto Destino Numero Factura Vr. Pago",
            "Empresa: UNION TEMPORAL ITALCO",
            "",
            None,
        ):
            self.assertIsNone(match_linea_transferencia(linea), msg=repr(linea))


class TestReconcileData(unittest.TestCase):
    def _despr(self, filas):
        return pd.DataFrame(filas)

    def test_cruce_por_documento_no_es_no_encontrada(self):
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3484422, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = conciliar(despr, trans, None)
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
        df_t, _ = conciliar(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "Valor no coincide")

    def test_sin_coincidencia_es_no_encontrada(self):
        despr = self._despr([
            {"Identificacion": "99999999", "Neto": 1000, "Devengado": None, "Cuenta": "111"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = conciliar(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "Transferencia no encontrada")

    def test_filtra_transferencias_de_otras_quincenas(self):
        # Caso real ORTEGON: desprendibles de abril (1Q + 2Q). Las transferencias
        # incluyen marzo (250331) y mayo (250515), que NO deben sumarse.
        despr = pd.DataFrame([
            {"Identificacion": "91513843", "Neto": 2447732, "Devengado": None,
             "Cuenta": "603168089", "PeriodoInicio": pd.Timestamp("2025-04-01"),
             "PeriodoFin": pd.Timestamp("2025-04-15")},
            {"Identificacion": "91513843", "Neto": 3457617, "Devengado": None,
             "Cuenta": "603168089", "PeriodoInicio": pd.Timestamp("2025-04-01"),
             "PeriodoFin": pd.Timestamp("2025-04-30")},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422, "Fecha": pd.Timestamp("2025-03-31")},
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 2447732, "Fecha": pd.Timestamp("2025-04-15")},
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3457617, "Fecha": pd.Timestamp("2025-04-30")},
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3008274, "Fecha": pd.Timestamp("2025-05-15")},
        ])
        df_t, _ = conciliar(despr, trans, None)
        fila = df_t.iloc[0]
        # Solo las dos de abril quedan; su suma coincide con los netos -> OK.
        self.assertEqual(fila["Estado"], "OK")
        self.assertEqual(sorted(fila["Valores_transferencia"]), [2447732, 3457617])

    def test_cruce_por_cuenta_con_ceros_de_relleno(self):
        # El documento difiere pero la cuenta coincide salvo ceros a la izquierda.
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3484422, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "0000000", "Cuenta": "00603168089", "Valor": 3484422},
        ])
        df_t, _ = conciliar(despr, trans, None)
        self.assertEqual(df_t.iloc[0]["Estado"], "OK")

    def test_candidata_se_rescata_por_valor_y_se_descarta_la_ajena(self):
        # Caso real SOTO JARABA: el pago 2Q viene en un layout sin etiqueta NÓMINA
        # ni fecha de quincena (candidato). Debe rescatarse porque su valor coincide
        # con el neto 2Q; un importe que no esté en los netos (p. ej. prima) se descarta.
        despr = pd.DataFrame([
            {"Identificacion": "1005179167", "Neto": 4457982, "Devengado": None, "Cuenta": "x",
             "PeriodoInicio": pd.Timestamp("2025-06-01"), "PeriodoFin": pd.Timestamp("2025-06-15")},
            {"Identificacion": "1005179167", "Neto": 2677442, "Devengado": None, "Cuenta": "x",
             "PeriodoInicio": pd.Timestamp("2025-06-01"), "PeriodoFin": pd.Timestamp("2025-06-30")},
        ])
        trans = pd.DataFrame([
            # Confiable (NÓMINA) con fecha dentro de la ventana.
            {"Documento": "1005179167", "Cuenta": "x", "Valor": 4457982,
             "Fecha": pd.Timestamp("2025-06-15"), "EsNomina": True},
            # Candidata sin fecha cuyo valor coincide con el neto 2Q -> se rescata.
            {"Documento": "1005179167", "Cuenta": "x", "Valor": 2677442,
             "Fecha": pd.NaT, "EsNomina": False},
            # Candidata cuyo valor NO está en los netos -> se descarta.
            {"Documento": "1005179167", "Cuenta": "x", "Valor": 4787967,
             "Fecha": pd.NaT, "EsNomina": False},
        ])
        df_t, _ = conciliar(despr, trans, None)
        fila = df_t.iloc[0]
        self.assertEqual(fila["Estado"], "OK")
        self.assertEqual(sorted(fila["Valores_transferencia"]), [2677442, 4457982])

    def test_transferencia_incluye_columna_diferencia(self):
        despr = self._despr([
            {"Identificacion": "91513843", "Neto": 3457617, "Devengado": None, "Cuenta": "603168089"},
        ])
        trans = pd.DataFrame([
            {"Documento": "91513843", "Cuenta": "603168089", "Valor": 3484422},
        ])
        df_t, _ = conciliar(despr, trans, None)
        fila = df_t.iloc[0]
        self.assertEqual(fila["Estado"], "Valor no coincide")
        self.assertEqual(fila["Diferencia"], 3457617 - 3484422)


class TestSeguridadSocial(unittest.TestCase):
    """Seguridad social: suma de IBC, devengados sin deduplicar y columna Diferencia."""

    def test_devengado_no_se_deduplica_y_suma_ibc_da_ok(self):
        # Dos quincenas con el MISMO devengado deben sumarse ambas (no deduplicar).
        despr = pd.DataFrame([
            {"Identificacion": "63472356", "Neto": 1, "Devengado": 2500950, "Cuenta": "x"},
            {"Identificacion": "63472356", "Neto": 1, "Devengado": 2500950, "Cuenta": "x"},
        ])
        seg = pd.DataFrame([
            {"cc": "63472356", "ibc": 2500950},
            {"cc": "63472356", "ibc": 2500950},
        ])
        _, df_s = conciliar(despr, None, seg)
        fila = df_s.iloc[0]
        self.assertEqual(fila["Devengado"], 5001900)   # ambos devengados sumados
        self.assertEqual(fila["Estado"], "OK")          # 5001900 == 2500950+2500950
        self.assertEqual(fila["Diferencia"], 0)

    def test_ok_por_suma_de_varios_ibc(self):
        # Varios IBC que sumados igualan el devengado -> OK (no por pertenencia).
        despr = pd.DataFrame([
            {"Identificacion": "1096185839", "Neto": 1, "Devengado": 8456828, "Cuenta": "x"},
        ])
        seg = pd.DataFrame([
            {"cc": "1096185839", "ibc": 182210},
            {"cc": "1096185839", "ibc": 8274618},
        ])
        _, df_s = conciliar(despr, None, seg)
        fila = df_s.iloc[0]
        self.assertEqual(fila["Estado"], "OK")
        self.assertEqual(fila["Diferencia"], 0)

    def test_diferencia_cuando_no_coincide(self):
        despr = pd.DataFrame([
            {"Identificacion": "111", "Neto": 1, "Devengado": 5001900, "Cuenta": "x"},
        ])
        seg = pd.DataFrame([{"cc": "111", "ibc": 5040619}])
        _, df_s = conciliar(despr, None, seg)
        fila = df_s.iloc[0]
        self.assertEqual(fila["Estado"], "Devengado no coincide")
        self.assertEqual(fila["Diferencia"], 5001900 - 5040619)


if __name__ == "__main__":
    unittest.main(verbosity=2)
