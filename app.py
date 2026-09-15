#!/usr/bin/env python3
"""
Extrae el historial de cambios de etapa de los leads de Kommo (API v4)
y genera un Excel con:
  - Hoja "HISTORIAL": un renglon por cada cambio de etapa
  - Hoja "PIVOTE":    un renglon por lead, una columna por etapa
                      con la PRIMERA fecha en que el lead entro a esa etapa

Requisitos:
    pip install requests pandas openpyxl

Uso:
    export KOMMO_TOKEN="eyJ0eXAiOi..."      # token de larga duracion
    python kommo_historial_etapas.py
"""

import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta
from threading import Lock
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import requests
import pandas as pd
import json

# ----------------------------------------------------------------------
# CONFIGURACION
# ----------------------------------------------------------------------

config = json.load(open("configuration.json"))
secret = json.load(open("secret.json"))


SUBDOMAIN = secret["kommo"]["SUBDOMAIN"]
TOKEN = os.environ.get("KOMMO_TOKEN", secret["kommo"]["TOKEN"])

# Filtra solo un embudo (None = todos los embudos)
PIPELINE_ID = secret["kommo"]["PIPELINE_ID"]

# Rango de fechas. FECHA_DESDE es necesaria para poder paralelizar.
FECHA_DESDE = config["settings"]["FECHA_DESDE"]   # formato YYYY-MM-DD
FECHA_HASTA = config["settings"]["FECHA_HASTA"]           # None = hasta hoy

# Incluir el evento de creacion del lead (da la etapa de entrada)
INCLUIR_LEAD_ADDED = config["settings"]["INCLUIR_LEAD_ADDED"]

# ---- Rendimiento y ruido -------------------------------------------------
VERBOSE = config["settings"]["performance"]["VERBOSE"]
HILOS = config["settings"]["performance"]["THREADS"]
INCLUIR_USUARIOS = config["settings"]["performance"]["INCLUIR_USUARIOS"]

# Zona horaria para mostrar las fechas (Mexico centro = -6)
TZ = timezone(timedelta(hours=-6))

SALIDA = "historial_etapas_kommo.xlsx"

BASE = f"https://{SUBDOMAIN}.kommo.com/api/v4"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

# Una sola sesion = una sola negociacion TLS reutilizada en todas las llamadas
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

_retry = Retry(total=5, connect=5, read=5, backoff_factor=1.5,
               status_forcelist=[429, 500, 502, 503, 504],
               allowed_methods=["GET"])
SESSION.mount("https://", HTTPAdapter(max_retries=_retry,
                                      pool_connections=HILOS,
                                      pool_maxsize=HILOS * 2))
SESSION.headers.update({"User-Agent": "kommo-export/1.0"})

def log(msg):
    if VERBOSE:
        print(msg, file=sys.stderr)

_prog = {"req": 0, "ev": 0}
_lock = Lock()

def avance(n_eventos):
    with _lock:
        _prog["ev"] += n_eventos
        print(f"\r {_prog['ev']} eventos",
              end="", file=sys.stderr, flush=True)

def limpiar_avance():
    print("\r" + " " * 40 + "\r", end="", file=sys.stderr, flush=True)

# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def a_timestamp(fecha_str):
    if not fecha_str:
        return None
    return int(datetime.strptime(fecha_str, "%Y-%m-%d")
               .replace(tzinfo=TZ).timestamp())


def get(url, params=None):
    ultimo_error = None
    for intento in range(5):
        try:
            r = SESSION.get(url, params=params, timeout=60)
        except requests.exceptions.RequestException as e:
            ultimo_error = e
            time.sleep(2 * (intento + 1))
            continue
        if r.status_code == 204:
            return None
        if r.status_code == 429 or r.status_code >= 500:
            ultimo_error = f"HTTP {r.status_code}"
            time.sleep(2 * (intento + 1))
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError(f"Fallaron 5 intentos en {url} — {ultimo_error}")


def cargar_catalogo_etapas():
    """Devuelve {status_id: (embudo, etapa, orden)}."""
    data = get(f"{BASE}/leads/pipelines")
    mapa = {}
    for pipe in data["_embedded"]["pipelines"]:
        for st in pipe["_embedded"]["statuses"]:
            mapa[st["id"]] = (pipe["name"], st["name"], st["sort"])
    return mapa


def cargar_usuarios():
    """Devuelve {user_id: nombre}. Se omite si INCLUIR_USUARIOS es False."""
    if not INCLUIR_USUARIOS:
        return {}
    mapa = {0: "Sistema / Bot"}
    page = 1
    while True:
        data = get(f"{BASE}/users", {"page": page, "limit": 250})
        if not data:
            break
        for u in data["_embedded"]["users"]:
            mapa[u["id"]] = u["name"]
        if not data.get("_links", {}).get("next"):
            break
        page += 1
    return mapa


def _params_base():
    tipos = ["lead_status_changed"]
    if INCLUIR_LEAD_ADDED:
        tipos.append("lead_added")
    p = {"filter[entity][]": "lead", "limit": 100}
    for i, t in enumerate(tipos):
        p[f"filter[type][{i}]"] = t
    return p


def _descargar_rango(desde, hasta):
    """Pagina un solo tramo de fechas."""
    params = _params_base()
    params["page"] = 1
    if desde:
        params["filter[created_at][from]"] = desde
    if hasta:
        params["filter[created_at][to]"] = hasta

    out = []
    while True:
        data = get(f"{BASE}/events", params)
        if not data:
            break
        out.extend(data["_embedded"]["events"])
        avance(len(data["_embedded"]["events"]))
        if not data.get("_links", {}).get("next"):
            break
        params["page"] += 1
    log(f"  tramo {desde}-{hasta}: {len(out)} eventos")
    return out


def descargar_eventos():
    """Parte el rango en HILOS tramos y los baja en paralelo."""
    desde, hasta = a_timestamp(FECHA_DESDE), a_timestamp(FECHA_HASTA)

    if desde is None or HILOS <= 1:
        return _descargar_rango(desde, hasta)

    fin = hasta or int(time.time())
    paso = max(1, (fin - desde) // HILOS)
    rangos, cursor = [], desde
    while cursor < fin:
        rangos.append((cursor, min(cursor + paso - 1, fin)))
        cursor += paso

    eventos = []
    with ThreadPoolExecutor(max_workers=HILOS) as ex:
        for lote in ex.map(lambda r: _descargar_rango(*r), rangos):
            eventos.extend(lote)

    # Dedup por si un evento cae justo en la frontera de dos tramos
    vistos, unicos = set(), []
    for ev in eventos:
        if ev["id"] not in vistos:
            vistos.add(ev["id"])
            unicos.append(ev)
    return unicos


def extraer_status(bloque):
    """Saca (status_id, pipeline_id) de value_before / value_after."""
    if not bloque:
        return None, None
    item = bloque[0] if isinstance(bloque, list) else bloque
    st = item.get("lead_status") or {}
    return st.get("id"), st.get("pipeline_id")


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    if not TOKEN:
        sys.exit("ERROR: falta la variable de entorno KOMMO_TOKEN")

    inicio = time.time()

    log("Cargando catalogo de embudos y etapas...")
    etapas = cargar_catalogo_etapas()
    usuarios = cargar_usuarios()

    log("Descargando eventos...")
    eventos = descargar_eventos()
    if not eventos:
        sys.exit("No se encontraron eventos en el rango solicitado.")

    filas = []
    for ev in eventos:
        sid_ant, pid_ant = extraer_status(ev.get("value_before"))
        sid_new, pid_new = extraer_status(ev.get("value_after"))
        pipeline_id = pid_new or pid_ant

        if PIPELINE_ID and pipeline_id != PIPELINE_ID:
            continue

        emb_new, eta_new, orden = etapas.get(sid_new, ("", "", 999))
        _, eta_ant, _ = etapas.get(sid_ant, ("", "", 999))

        fecha = datetime.fromtimestamp(ev["created_at"], TZ)

        fila = {
            "LEAD_ID": ev["entity_id"],
            "EMBUDO": emb_new,
            "ETAPA_ANTERIOR": eta_ant,
            "ETAPA_NUEVA": eta_new,
            "FECHA": fecha.replace(tzinfo=None),
            "TIPO_EVENTO": ev["type"],
            "ORDEN_ETAPA": orden,
        }
        if INCLUIR_USUARIOS:
            fila["MOVIDO_POR"] = usuarios.get(ev.get("created_by"),
                                              ev.get("created_by"))
        filas.append(fila)

    df = pd.DataFrame(filas).sort_values(["LEAD_ID", "FECHA"])

    # Dias entre un movimiento y el siguiente del mismo lead
    df["DIAS_EN_ETAPA_ANTERIOR"] = (
        df.groupby("LEAD_ID")["FECHA"].diff().dt.total_seconds() / 86400
    ).round(2)

    # Pivote: primera fecha en que cada lead toco cada etapa
    orden_cols = (df[["ETAPA_NUEVA", "ORDEN_ETAPA"]]
                  .drop_duplicates()
                  .sort_values("ORDEN_ETAPA")["ETAPA_NUEVA"].tolist())

    pivote = (df.pivot_table(index="LEAD_ID",
                             columns="ETAPA_NUEVA",
                             values="FECHA",
                             aggfunc="min")
                .reindex(columns=orden_cols)
                .reset_index())

    with pd.ExcelWriter(SALIDA, engine="openpyxl",
                        datetime_format="yyyy-mm-dd hh:mm:ss") as xl:
        df.drop(columns=["ORDEN_ETAPA"]).to_excel(
            xl, sheet_name="HISTORIAL", index=False)
        pivote.to_excel(xl, sheet_name="PIVOTE", index=False)

        for hoja, ancho in (("HISTORIAL", 22), ("PIVOTE", 20)):
            ws = xl.sheets[hoja]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = ancho

    limpiar_avance()
    print(f"{SALIDA} | {len(df)} movimientos | "
          f"{df['LEAD_ID'].nunique()} leads | {time.time() - inicio:.1f}s")


if __name__ == "__main__":
    main()