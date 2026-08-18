"""Perfiles TABARCA: respaldo por Cargo cuando el nombre no cabe en Nivel/Perfil.

Regresion del caso real: un perfil de nombre largo llega partido en la columna
``Nivel/Perfil`` del PDF (``Inspector certificad o:``) mientras el nombre
completo si esta en ``Cargo``. Antes salian dos filas descuadradas; ahora deben
cruzar como una sola.

    python -m unittest discover -s web_ui -p "test_*.py" -t .
"""
import os
import sys
import unittest

import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
sys.path.insert(0, os.path.join(RAIZ, "web_ui"))

from processing import _aplicar_respaldo_cargo, _indice_columna_cargo  # noqa: E402

# Encabezado real de la planilla TABARCA: ocupa DOS filas. "Nivel/ Perfil" y la
# fecha van en la 6; "CEDULA | TRABAJADOR | CARGO" en la 7; los datos en la 8.
PLANILLA = [
    ["1.INFORMACION GENERAL", "", "", "", "", "", "", "", "", "", ""],
    ["EMPRESA:", "", "CONSORCIO TABARCA", "", "", "", "", "", "NIT:", "901290945-6", ""],
    ["DIRECCION:", "", "CARRERA 30", "", "", "", "", "", "", "", ""],
    ["OBJETO:", "", "OSM064", "", "", "", "", "", "", "", ""],
    # Ojo: esta fila trae un "CARGO:" que es el del responsable, no la columna.
    ["RESPONSABLE:", "", "MARLON PAJARO", "", "", "", "", "", "CARGO:", "DIRECTOR", ""],
    ["TIPO ORDEN:", "", "Ejecucion", "", "", "", "", "", "", "", ""],
    ["DATOS", "", "", "Nivel/ Perfil", "C O D I", "HORARIO", "", "H O R", "FECHA:", "", "20-07-2026"],
    ["CEDULA", "TRABAJADOR", "CARGO", "", "G O", "INGRESO", "SALIDA", "A S", "CUADRILLA", "FIRMA", "OBSERVACIONES"],
    ["80849778", "JONATHAN HERNANDEZ", "AYUDANTE TECNICO", "B4", "M11", "07:00", "15:00", "8,0", "---", "", ""],
]


class TestIndiceColumnaCargo(unittest.TestCase):
    def test_encuentra_cargo_en_la_segunda_fila_de_encabezado(self):
        self.assertEqual(_indice_columna_cargo(PLANILLA), 2)

    def test_no_confunde_el_cargo_del_responsable(self):
        # Sin la fila 7, el unico "CARGO:" es el del responsable (fila 4): no
        # debe tomarse como columna de la tabla.
        sin_subcabecera = [f for i, f in enumerate(PLANILLA) if i != 7]
        self.assertIsNone(_indice_columna_cargo(sin_subcabecera))

    def test_tabla_corta_no_falla(self):
        self.assertIsNone(_indice_columna_cargo(PLANILLA[:3]))

FECHA = pd.Timestamp("2026-07-22")
LARGO = "Inspector certificado: APIASME NACIONAL"
PARTIDO = "Inspector certificad o:"


def _pdf(filas):
    return pd.DataFrame(
        [
            {
                "FECHA": FECHA,
                "PERFIL_NORM": perfil,
                "Nivel/Perfil": perfil,
                "PDF": pdf,
                "CARGO_NORM": cargo,
            }
            for perfil, cargo, pdf in filas
        ]
    )


def _excel(filas):
    return pd.DataFrame(
        [
            {"FECHA": FECHA, "PERFIL_NORM": p, "Nivel/Perfil": p, "Excel": v}
            for p, v in filas
        ]
    )


class TestRespaldoPorCargo(unittest.TestCase):
    def test_perfil_partido_cruza_por_el_cargo(self):
        df = _aplicar_respaldo_cargo(
            _pdf([(PARTIDO, LARGO, 1)]), _excel([(LARGO, 1)])
        )
        self.assertEqual(list(df["PERFIL_NORM"]), [LARGO])
        self.assertEqual(list(df["PDF"]), [1])

    def test_perfil_que_ya_cruzaba_no_se_toca(self):
        # A2 esta en el Excel: aunque tenga un Cargo distinto, se respeta.
        df = _aplicar_respaldo_cargo(
            _pdf([("A2", "Almacenista", 3)]), _excel([("A2", 3), ("Almacenista", 9)])
        )
        self.assertEqual(list(df["PERFIL_NORM"]), ["A2"])
        self.assertEqual(list(df["PDF"]), [3])

    def test_cargo_que_no_esta_en_el_excel_no_cambia_nada(self):
        df = _aplicar_respaldo_cargo(
            _pdf([("X9", "Cargo inventado", 1)]), _excel([("A2", 1)])
        )
        self.assertEqual(list(df["PERFIL_NORM"]), ["X9"])

    def test_coincidencia_ignora_mayusculas_y_espacios_repetidos(self):
        df = _aplicar_respaldo_cargo(
            _pdf([(PARTIDO, "inspector  certificado:   APIASME nacional", 1)]),
            _excel([(LARGO, 1)]),
        )
        self.assertEqual(list(df["PERFIL_NORM"]), [LARGO])

    def test_no_pisa_un_perfil_que_el_pdf_ya_reporta(self):
        # Si el PDF ya trae una fila con ese nombre, reasignar duplicaria el
        # conteo: se deja como estaba.
        df = _aplicar_respaldo_cargo(
            _pdf([(PARTIDO, LARGO, 1), (LARGO, LARGO, 2)]), _excel([(LARGO, 3)])
        )
        self.assertEqual(sorted(df["PERFIL_NORM"]), sorted([PARTIDO, LARGO]))

    def test_suma_cuando_varias_filas_van_al_mismo_nombre(self):
        df = _aplicar_respaldo_cargo(
            _pdf([(PARTIDO, LARGO, 1), ("Inspector certific ado:", LARGO, 2)]),
            _excel([(LARGO, 3)]),
        )
        self.assertEqual(list(df["PERFIL_NORM"]), [LARGO])
        self.assertEqual(list(df["PDF"]), [3])

    def test_sin_columna_de_cargo_no_falla(self):
        sin_cargo = _pdf([(PARTIDO, None, 1)]).drop(columns=["CARGO_NORM"])
        df = _aplicar_respaldo_cargo(sin_cargo, _excel([(LARGO, 1)]))
        self.assertEqual(list(df["PERFIL_NORM"]), [PARTIDO])


if __name__ == "__main__":
    unittest.main(verbosity=2)
