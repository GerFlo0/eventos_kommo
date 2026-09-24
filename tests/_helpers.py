"""Utilidades comunes de las pruebas (sin red y sin secret.json real)."""
import copy
import importlib
import os
import sys
from pathlib import Path
from unittest import mock

TESTS = Path(__file__).resolve().parent
ROOT = TESTS.parent
for ruta in (str(ROOT), str(TESTS)):
    if ruta not in sys.path:
        sys.path.insert(0, ruta)

import fake_kommo  # noqa: E402
import functions as fn  # noqa: E402

ASESORES = ["ASESOR UNO CIERRES", "ANA PRUEBA CIERRES", "JULIANA PRUEBA CIERRES"]

SECRET = {
    "kommo": {"SUBDOMAIN": "demo", "TOKEN": "tok-secret",
              "PIPELINE_ID": {"CIERRES": fake_kommo.CIERRES, "VENTAS": fake_kommo.VENTAS}},
    "asesores": ASESORES,
    "query": "SELECT * FROM df WHERE LEAD_ID IN "
             "(SELECT LEAD_ID FROM df WHERE ETIQUETAS ILIKE $1)",
}

CONFIG = {"settings": {
    "FECHA_DESDE": "2026-09-01",
    "FECHA_HASTA": "2026-09-22",
    "INCLUIR_LEAD_ADDED": True,
    "SALIDA": "tablas/historial_etapas_kommo.xlsx",
    "ARCHIVOS_SEPARADOS": False,
    "FILTRO_POR_ETAPA_ACTUAL": True,
    "CAMPOS_TARJETA": ["ESTATUS DE NEGOCIO", "FECHA DICTAMEN"],
    "CAMPOS_REQUERIDOS": ["ESTATUS DE NEGOCIO", "FECHA DICTAMEN"],
    "EMBUDOS": [
        {"NOMBRE": "CIERRES", "CLAVE_SECRET": "CIERRES", "TODAS_LAS_ETAPAS": True},
        {"NOMBRE": "VENTAS", "CLAVE_SECRET": "VENTAS", "TODAS_LAS_ETAPAS": False,
         "ETAPAS_EXACTAS": ["SIN CAPACIDAD"], "ETAPAS_CONTIENEN": ["20"]},
    ],
    "performance": {"VERBOSE": False, "THREADS": 4, "INCLUIR_USUARIOS": False},
}}


def cargar_extractor(config=None, secret=None, token_env=None):
    """Importa extract_data_from_kommo desde cero con config/secret falsos.

    token_env: valor para la variable KOMMO_TOKEN (None = sin variable).
    """
    config = copy.deepcopy(config or CONFIG)
    secret = copy.deepcopy(secret or SECRET)
    real = fn.import_json

    def falso(ruta):
        ruta = str(ruta)
        if ruta.endswith("secret.json"):
            return copy.deepcopy(secret)
        if ruta.endswith("configuration.json"):
            return copy.deepcopy(config)
        return real(ruta)

    entorno = {k: v for k, v in os.environ.items() if k != "KOMMO_TOKEN"}
    if token_env is not None:
        entorno["KOMMO_TOKEN"] = token_env
    sys.modules.pop("extract_data_from_kommo", None)
    with mock.patch.object(fn, "import_json", falso), \
            mock.patch.dict(os.environ, entorno, clear=True):
        return importlib.import_module("extract_data_from_kommo")
