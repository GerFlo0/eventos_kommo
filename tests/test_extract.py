import contextlib
import copy
import io
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import pandas as pd
import requests

from _helpers import CONFIG, SECRET, cargar_extractor, fake_kommo

C, V = fake_kommo.CIERRES, fake_kommo.VENTAS


def respuesta(status, data=None, headers=None):
    return fake_kommo.FakeResponse(status, data, headers)


class TestFechas(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ex = cargar_extractor()

    def test_desde_es_inicio_del_dia(self):
        t = self.ex.a_timestamp("2026-09-15")
        self.assertEqual(datetime.fromtimestamp(t, self.ex.TZ).strftime("%Y-%m-%d %H:%M:%S"),
                         "2026-09-15 00:00:00")

    def test_hasta_incluye_el_dia_completo(self):
        t = self.ex.a_timestamp("2026-09-15", fin_de_dia=True)
        self.assertEqual(datetime.fromtimestamp(t, self.ex.TZ).strftime("%Y-%m-%d %H:%M:%S"),
                         "2026-09-15 23:59:59")

    def test_vacio(self):
        self.assertIsNone(self.ex.a_timestamp(None))
        self.assertIsNone(self.ex.a_timestamp("", fin_de_dia=True))


class TestCatalogo(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ex = cargar_extractor()
        fake = fake_kommo.FakeKommo(fake_kommo.NOMBRES_DISTINTOS)
        with mock.patch.object(cls.ex.SESSION, "get", fake.get):
            cls.etapas, cls.nombres = cls.ex.cargar_catalogo_etapas()

    def test_142_y_143_no_se_pisan_entre_embudos(self):
        self.assertEqual(self.etapas[(C, 142)][2], "Logrado con éxito")
        self.assertEqual(self.etapas[(V, 142)][2], "Ganado ventas")
        self.assertEqual(self.etapas[(C, 143)][2], "Venta perdida")
        self.assertEqual(self.etapas[(V, 143)][2], "Ventas perdidas")

    def test_buscar_etapa(self):
        b = self.ex.buscar_etapa
        self.assertEqual(b(self.etapas, C, 142)[2], "Logrado con éxito")
        self.assertEqual(b(self.etapas, V, 23)[2], "SIN CAPACIDAD")
        # evento sin pipeline_id: cae a buscar solo por status_id
        self.assertEqual(b(self.etapas, None, 23)[2], "SIN CAPACIDAD")
        self.assertEqual(b(self.etapas, C, 99999), self.ex.SIN_ETAPA)
        self.assertEqual(b(self.etapas, None, None), self.ex.SIN_ETAPA)

    def test_buscar_etapa_con_catalogo_de_llave_antigua(self):
        antiguo = {sid: v for (_pid, sid), v in self.etapas.items()}
        self.assertEqual(self.ex.buscar_etapa(antiguo, V, 23)[2], "SIN CAPACIDAD")

    def test_resolver_embudos_filtra_etapas_de_ventas(self):
        emb = self.ex.resolver_embudos(self.etapas, self.nombres)
        self.assertEqual(list(emb), [C, V])
        self.assertTrue(emb[C]["todas"] and emb[C]["principal"])
        self.assertEqual(emb[V]["etapas_ids"], {23, 24, 25})

    def test_resolver_embudos_con_catalogo_de_llave_antigua(self):
        antiguo = {sid: v for (_pid, sid), v in self.etapas.items() if _pid != 3003}
        emb = self.ex.resolver_embudos(antiguo, self.nombres)
        self.assertEqual(emb[V]["etapas_ids"], {23, 24, 25})

    def test_etapa_de_sistema_se_encuentra_en_su_embudo(self):
        config = copy.deepcopy(CONFIG)
        config["settings"]["EMBUDOS"][1]["ETAPAS_EXACTAS"] = ["Ventas perdidas"]
        config["settings"]["EMBUDOS"][1]["ETAPAS_CONTIENEN"] = []
        ex = cargar_extractor(config)
        emb = ex.resolver_embudos(self.etapas, self.nombres)
        self.assertEqual(emb[V]["etapas_ids"], {143})


class TestGet(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ex = cargar_extractor()

    def correr(self, respuestas):
        """Ejecuta get() con una lista de respuestas/excepciones simuladas."""
        sesion = mock.Mock(side_effect=respuestas)
        with mock.patch.object(self.ex.SESSION, "get", sesion), \
                mock.patch.object(self.ex.time, "sleep") as dormir:
            try:
                return self.ex.get("https://x/api/v4/events"), sesion, dormir
            except Exception as e:  # noqa: BLE001
                return e, sesion, dormir

    def test_reintenta_un_429_y_devuelve_datos(self):
        r, sesion, dormir = self.correr([respuesta(429), respuesta(200, {"ok": 1})])
        self.assertEqual(r, {"ok": 1})
        self.assertEqual(sesion.call_count, 2)
        dormir.assert_called_once_with(2)

    def test_error_persistente_falla_tras_5_intentos(self):
        r, sesion, dormir = self.correr([respuesta(503)] * 5)
        self.assertIsInstance(r, RuntimeError)
        self.assertIn("5 intentos", str(r))
        self.assertEqual(sesion.call_count, 5)
        self.assertEqual([c.args[0] for c in dormir.call_args_list], [2, 4, 8, 16])

    def test_respeta_retry_after(self):
        _, _, dormir = self.correr([respuesta(429, headers={"Retry-After": "10"}),
                                    respuesta(200, {})])
        dormir.assert_called_once_with(10.0)

    def test_retry_after_tiene_tope(self):
        _, _, dormir = self.correr([respuesta(429, headers={"Retry-After": "999"}),
                                    respuesta(200, {})])
        dormir.assert_called_once_with(60)

    def test_204_devuelve_none(self):
        r, sesion, _ = self.correr([respuesta(204)])
        self.assertIsNone(r)

    def test_404_falla_de_inmediato(self):
        r, sesion, dormir = self.correr([respuesta(404)])
        self.assertIsInstance(r, requests.exceptions.HTTPError)
        self.assertEqual(sesion.call_count, 1)
        dormir.assert_not_called()

    def test_error_de_conexion_se_reintenta(self):
        r, sesion, _ = self.correr([requests.exceptions.ConnectionError("x"),
                                    respuesta(200, {"ok": 2})])
        self.assertEqual(r, {"ok": 2})

    def test_adaptador_sin_reintentos_propios(self):
        adaptador = self.ex.SESSION.get_adapter("https://demo.kommo.com")
        self.assertEqual(adaptador.max_retries.total, 0)


class TestToken(unittest.TestCase):
    def test_variable_de_entorno_tiene_prioridad(self):
        self.assertEqual(cargar_extractor(token_env="tok-env").TOKEN, "tok-env")

    def test_sin_variable_usa_secret(self):
        self.assertEqual(cargar_extractor().TOKEN, "tok-secret")

    def test_secret_sin_token_no_truena_si_hay_variable(self):
        secret = copy.deepcopy(SECRET)
        del secret["kommo"]["TOKEN"]
        self.assertEqual(cargar_extractor(secret=secret, token_env="tok-env").TOKEN, "tok-env")

    def test_sin_token_main_termina_con_mensaje(self):
        secret = copy.deepcopy(SECRET)
        del secret["kommo"]["TOKEN"]
        ex = cargar_extractor(secret=secret)
        with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(SystemExit) as ctx:
            ex.main()
        self.assertIn("token", str(ctx.exception))


class TestFlujoCompleto(unittest.TestCase):
    def correr(self, config=None, nombres=fake_kommo.NOMBRES_IGUALES):
        ex = cargar_extractor(config)
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        ex.SALIDA = str(Path(tmp.name) / "sub" / "historial.xlsx")
        fake = fake_kommo.FakeKommo(nombres)
        salida = io.StringIO()
        with mock.patch.object(ex.SESSION, "get", fake.get), \
                mock.patch.object(ex.time, "sleep"), \
                contextlib.redirect_stdout(salida), contextlib.redirect_stderr(salida):
            ex.main()
        return Path(tmp.name), salida.getvalue()

    def leer(self, ruta):
        return pd.read_excel(ruta, sheet_name=None)

    def test_un_archivo(self):
        carpeta, log = self.correr()
        hojas = self.leer(carpeta / "sub" / "historial.xlsx")
        self.assertEqual(set(hojas), {"HISTORIAL", "PIVOTE"})
        h = hojas["HISTORIAL"]
        self.assertEqual(list(h.columns)[:5],
                         ["LEAD_ID", "EMBUDO", "ESTATUS DE NEGOCIO", "FECHA DICTAMEN", "ETIQUETAS"])
        # 107 no tiene FECHA DICTAMEN; 109 es de otro embudo; 106 solo entra por CIERRES
        self.assertNotIn(107, h["LEAD_ID"].values)
        self.assertNotIn(109, h["LEAD_ID"].values)
        self.assertEqual(set(h.loc[h["EMBUDO"] == "VENTAS", "LEAD_ID"]), {103, 104, 108, 113})
        # las etiquetas solo van en la fila más reciente de cada lead
        self.assertEqual(h["ETIQUETAS"].notna().sum(), h.loc[h["ETIQUETAS"].notna(), "LEAD_ID"].nunique())
        self.assertIn("movimientos", log)

    def test_142_con_nombre_de_su_embudo(self):
        carpeta, _ = self.correr(nombres=fake_kommo.NOMBRES_DISTINTOS)
        h = self.leer(carpeta / "sub" / "historial.xlsx")["HISTORIAL"]
        fila = h[(h["LEAD_ID"] == 101)].iloc[-1]
        self.assertEqual(fila["ETAPA_NUEVA"], "Logrado con éxito")

    def test_fecha_hasta_incluye_ese_dia(self):
        config = copy.deepcopy(CONFIG)
        config["settings"]["FECHA_HASTA"] = "2026-09-15"
        carpeta, _ = self.correr(config)
        h = self.leer(carpeta / "sub" / "historial.xlsx")["HISTORIAL"]
        fechas_101 = pd.to_datetime(h.loc[h["LEAD_ID"] == 101, "FECHA"])
        self.assertEqual(fechas_101.max(), pd.Timestamp("2026-09-15 11:30"))
        self.assertTrue((pd.to_datetime(h["FECHA"]) < pd.Timestamp("2026-09-16")).all())

    def test_archivos_separados(self):
        config = copy.deepcopy(CONFIG)
        config["settings"]["ARCHIVOS_SEPARADOS"] = True
        config["settings"]["performance"]["INCLUIR_USUARIOS"] = True
        carpeta, _ = self.correr(config)
        c = self.leer(carpeta / "sub" / "historial_CIERRES.xlsx")["HISTORIAL"]
        v = self.leer(carpeta / "sub" / "historial_VENTAS.xlsx")["HISTORIAL"]
        self.assertEqual(set(c["EMBUDO"]), {"CIERRES"})
        self.assertEqual(set(v["EMBUDO"]), {"VENTAS"})
        self.assertIn("MOVIDO_POR", c.columns)


if __name__ == "__main__":
    unittest.main()
