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
    header = [f"c{i}" for i in range(n_cols)]             # fila 9 (encabezado)
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
