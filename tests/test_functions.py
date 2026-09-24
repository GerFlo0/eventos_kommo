import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

from _helpers import ROOT, fn


class TestRutas(unittest.TestCase):
    def test_ruta_proyecto_relativa_y_absoluta(self):
        self.assertEqual(fn.ruta_proyecto("json/x.json"), ROOT / "json/x.json")
        absoluta = Path(tempfile.gettempdir()) / "x.json"
        self.assertEqual(fn.ruta_proyecto(absoluta), absoluta)

    def test_lectura_funciona_desde_otra_carpeta(self):
        actual = os.getcwd()
        try:
            os.chdir(tempfile.gettempdir())
            config = fn.import_json("json/configuration.json")
        finally:
            os.chdir(actual)
        self.assertIn("settings", config)

    def test_lectura_prefiere_la_carpeta_actual(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "json").mkdir()
            (Path(tmp) / "json/configuration.json").write_text('{"local": true}')
            actual = os.getcwd()
            try:
                os.chdir(tmp)
                self.assertEqual(fn.import_json("json/configuration.json"), {"local": True})
            finally:
                os.chdir(actual)


class TestImportJson(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_utf8_con_y_sin_bom(self):
        for cod in ("utf-8", "utf-8-sig"):
            ruta = self.dir / f"{cod}.json"
            ruta.write_text('{"n": "Gómez"}', encoding=cod)
            self.assertEqual(fn.import_json(ruta), {"n": "Gómez"})

    def test_ansi_de_windows_sigue_funcionando(self):
        ruta = self.dir / "ansi.json"
        ruta.write_text('{"n": "Gómez"}', encoding="cp1252")
        with mock.patch("locale.getpreferredencoding", return_value="cp1252"):
            self.assertEqual(fn.import_json(ruta), {"n": "Gómez"})

    def test_archivo_inexistente_da_error_claro(self):
        with self.assertRaises(FileNotFoundError) as ctx:
            fn.import_json(self.dir / "no_existe.json")
        self.assertIn("no_existe.json", str(ctx.exception))

    def test_json_invalido_indica_linea(self):
        ruta = self.dir / "malo.json"
        ruta.write_text('{\n  "a": 1,\n}')
        with self.assertRaises(ValueError) as ctx:
            fn.import_json(ruta)
        self.assertIn("línea 3", str(ctx.exception))


class TestImportXlsx(unittest.TestCase):
    def test_lee_excel(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "a.xlsx"
            pd.DataFrame({"A": [1, 2]}).to_excel(ruta, index=False)
            self.assertEqual(fn.import_xlsx(ruta)["A"].tolist(), [1, 2])

    def test_archivo_inexistente_da_error_claro(self):
        with self.assertRaises(FileNotFoundError):
            fn.import_xlsx("tablas/no_existe_nunca.xlsx")


if __name__ == "__main__":
    unittest.main()
