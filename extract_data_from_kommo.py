#!/usr/bin/env python3
"""
Extrae el historial de cambios de etapa de los leads de Kommo (API v4)
y genera un Excel con:
  - Hoja "HISTORIAL": un renglon por cada cambio de etapa
                      (FECHA lleva fecha y hora en la misma celda)
  - Hoja "PIVOTE":    un renglon por lead, una columna por etapa
                      con la PRIMERA fecha en que el lead entro a esa etapa

Trabaja con VARIOS EMBUDOS a la vez (por defecto CIERRES + VENTAS).
De cada embudo se puede traer todo o solo las etapas que interesen.

Requisitos:
    Ejecutar antes setup_invironment.py (instala requirements.txt).

Uso:
    python extract_data_from_kommo.py

    El token se toma de la variable de entorno KOMMO_TOKEN si existe;
    si no, de secret.json -> kommo -> TOKEN.

----------------------------------------------------------------------
COMO CAMBIAR ENTRE "UN SOLO ARCHIVO" Y "UN ARCHIVO POR EMBUDO"
----------------------------------------------------------------------
En configuration.json:

    "ARCHIVOS_SEPARADOS": false   ->  UN solo Excel con todo junto
                                      (lo de VENTAS se agrega despues
                                       de lo de CIERRES).   <-- default
    "ARCHIVOS_SEPARADOS": true    ->  UN Excel por embudo:
                                      historial_etapas_CIERRES.xlsx
                                      historial_etapas_VENTAS.xlsx

Es lo unico que hay que tocar; se puede ir y venir las veces que sea.
Nota: la columna DIAS_EN_ETAPA_ANTERIOR se calcula con TODOS los
movimientos del lead que quedan en el reporte (aunque haya cruzado de un
embudo a otro), asi que da el mismo numero en los dos modos. Ojo: son los
movimientos que pasan los filtros y caen dentro del rango de fechas; si un
filtro descarta un movimiento intermedio, esos dias se suman al siguiente,
y el primer movimiento de cada lead dentro del rango queda vacio.
----------------------------------------------------------------------
"""

import os
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta
from threading import Lock

import requests
import pandas as pd
from requests.adapters import HTTPAdapter

import functions as fn

# ----------------------------------------------------------------------
# CONFIGURACION
# ----------------------------------------------------------------------

config = fn.import_json("json/configuration.json")
secret = fn.import_json("json/secret.json")

def a_bool(valor, por_defecto=False):
    """Acepta true/false reales y tambien los textos "True"/"False".

    (En JSON, "False" entre comillas es un texto y Python lo toma como
    verdadero; este helper evita esa trampa.)
    """
    if valor is None:
        return por_defecto
    if isinstance(valor, bool):
        return valor
    return str(valor).strip().lower() in ("true", "1", "si", "sí", "yes", "y")


SUBDOMAIN = secret["kommo"]["SUBDOMAIN"]
TOKEN = os.environ.get("KOMMO_TOKEN") or secret["kommo"].get("TOKEN")

# ---- Embudos a procesar --------------------------------------------------
# Se definen en configuration.json -> settings.EMBUDOS (lista ordenada).
# El primero de la lista es el embudo PRINCIPAL: manda en el orden de las
# filas, en el orden de las columnas del pivote y en el nombre del archivo.
EMBUDOS_CFG = config["settings"]["EMBUDOS"]

# Rango de fechas. FECHA_DESDE es necesaria para poder paralelizar.
FECHA_DESDE = config["settings"]["FECHA_DESDE"]   # formato YYYY-MM-DD
FECHA_HASTA = config["settings"]["FECHA_HASTA"]   # None = hasta hoy
                                                  # (el dia indicado se
                                                  # incluye completo)

# Incluir el evento de creacion del lead (da la etapa de entrada)
INCLUIR_LEAD_ADDED = a_bool(config["settings"]["INCLUIR_LEAD_ADDED"], True)

# Un solo archivo (False) o un archivo por embudo (True)
ARCHIVOS_SEPARADOS = a_bool(config["settings"].get("ARCHIVOS_SEPARADOS"), False)

# Como se decide que leads entran cuando un embudo esta filtrado por etapas:
#   True  -> entran los leads que HOY estan parados en esas etapas, y de
#            ellos se trae TODO su historial (recomendado / default).
#   False -> entran solo los movimientos cuya ETAPA_NUEVA es una de esas
#            etapas (no se consulta el estado actual del lead).
FILTRO_POR_ETAPA_ACTUAL = a_bool(
    config["settings"].get("FILTRO_POR_ETAPA_ACTUAL"), True)

# ---- Campos de la tarjeta del lead ---------------------------------------
# CAMPOS_TARJETA    = campos personalizados que se agregan como columnas
#                     en la hoja HISTORIAL (se escriben tal cual aparecen
#                     en Kommo; no importan acentos ni mayusculas).
# CAMPOS_REQUERIDOS = de esos campos, los que el lead DEBE tener llenos
#                     para aparecer en el reporte. Si un lead tiene vacio
#                     alguno, se descarta completo.
#                     Lista vacia [] = no se filtra por campos.
# Se puede sobreescribir por embudo: basta con poner CAMPOS_REQUERIDOS
# dentro del bloque del embudo en configuration.json (por ejemplo [] en
# VENTAS para exigirlos solo en CIERRES).
CAMPOS_TARJETA = list(config["settings"].get("CAMPOS_TARJETA", []))
CAMPOS_REQUERIDOS = list(config["settings"].get("CAMPOS_REQUERIDOS", []))

# Columna ETIQUETAS: las etiquetas (tags) del lead. Para no repetirlas en
# cada renglon, solo se escriben en la fila de la FECHA MAS RECIENTE de ese
# lead; si el lead tiene una sola fila, van en esa. Se apaga con false.
INCLUIR_ETIQUETAS = a_bool(config["settings"].get("INCLUIR_ETIQUETAS"), True)
COL_ETIQUETAS = "ETIQUETAS"

# ---- Rendimiento y ruido -------------------------------------------------
VERBOSE = a_bool(config["settings"]["performance"]["VERBOSE"], False)
HILOS = int(config["settings"]["performance"]["THREADS"])
INCLUIR_USUARIOS = a_bool(
    config["settings"]["performance"]["INCLUIR_USUARIOS"], False)

# Zona horaria para mostrar las fechas (Mexico centro = -6)
TZ = timezone(timedelta(hours=-6))

SALIDA = config["settings"].get("SALIDA", "historial_etapas_kommo.xlsx")
# Relativa a la raiz del proyecto, para que funcione aunque el script se
# ejecute desde otra carpeta.
SALIDA = str(fn.ruta_proyecto(SALIDA))

BASE = f"https://{SUBDOMAIN}.kommo.com/api/v4"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

# Una sola sesion = una sola negociacion TLS reutilizada en todas las llamadas
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# Los reintentos se hacen SOLO en get() (abajo). Antes tambien los hacia el
# adaptador HTTP y se multiplicaban: ante una caida de Kommo cada URL podia
# quedarse varios minutos reintentando antes de fallar.
SESSION.mount("https://", HTTPAdapter(pool_connections=HILOS,
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
def a_timestamp(fecha_str, fin_de_dia=False):
    """YYYY-MM-DD -> timestamp.

    fin_de_dia=False -> 00:00:00 de ese dia (para FECHA_DESDE).
    fin_de_dia=True  -> 23:59:59 de ese dia (para FECHA_HASTA), asi los
                        movimientos de ese dia tambien entran.
    """
    if not fecha_str:
        return None
    inicio = datetime.strptime(fecha_str, "%Y-%m-%d").replace(tzinfo=TZ)
    if fin_de_dia:
        inicio += timedelta(days=1, seconds=-1)
    return int(inicio.timestamp())


def normalizar(texto):
    """MAYUSCULAS, sin acentos y sin espacios de sobra, para comparar."""
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return " ".join(texto.upper().split())


INTENTOS = 5


def _espera(respuesta, intento):
    """Segundos a esperar antes del siguiente intento: 2, 4, 8, 16...

    Si Kommo manda el encabezado Retry-After (tipico en un 429), se
    respeta, con un maximo de 60 segundos.
    """
    espera = min(2 ** (intento + 1), 30)
    try:
        pedido = respuesta.headers.get("Retry-After")
        if pedido is not None:
            espera = min(max(float(pedido), espera), 60)
    except (AttributeError, TypeError, ValueError):
        pass
    return espera


def get(url, params=None):
    ultimo_error = None
    for intento in range(INTENTOS):
        r = None
        try:
            r = SESSION.get(url, params=params, timeout=60)
        except requests.exceptions.RequestException as e:
            ultimo_error = e
        else:
            if r.status_code == 204:
                return None
            if r.status_code == 429 or r.status_code >= 500:
                ultimo_error = f"HTTP {r.status_code}"
            else:
                r.raise_for_status()
                return r.json()
        if intento < INTENTOS - 1:          # no esperar despues del ultimo
            time.sleep(_espera(r, intento))
    raise RuntimeError(f"Fallaron {INTENTOS} intentos en {url} — {ultimo_error}")


# ----------------------------------------------------------------------
# CATALOGOS
# ----------------------------------------------------------------------
def cargar_catalogo_etapas():
    """Devuelve dos mapas:
       etapas  = {(pipeline_id, status_id): (pipeline_id, embudo, etapa, orden)}
       embudos = {pipeline_id: nombre_del_embudo}

    La llave incluye el embudo porque en Kommo los estados de sistema
    142 (Logrado con exito) y 143 (Venta perdida) tienen el MISMO id en
    todos los embudos; con solo el status_id, el ultimo embudo leido
    pisaba a los demas.
    """
    data = get(f"{BASE}/leads/pipelines")
    etapas, embudos = {}, {}
    for pipe in data["_embedded"]["pipelines"]:
        embudos[pipe["id"]] = pipe["name"]
        for st in pipe["_embedded"]["statuses"]:
            etapas[(pipe["id"], st["id"])] = (pipe["id"], pipe["name"],
                                              st["name"], st["sort"])
    return etapas, embudos


def _status_id(clave):
    """status_id de una llave del catalogo.

    Acepta la llave nueva (pipeline_id, status_id) y tambien la antigua
    (solo status_id), por compatibilidad.
    """
    return clave[1] if isinstance(clave, tuple) else clave


SIN_ETAPA = (None, "", "", 999)


def buscar_etapa(etapas, pipeline_id, status_id):
    """(pipeline_id, embudo, etapa, orden) de un estado del catalogo.

    Busca primero por (embudo, estado). Si el evento no trae el embudo, o
    el catalogo usa la llave antigua, cae a buscar solo por status_id
    (lo que hacia la version anterior). Si no existe, devuelve SIN_ETAPA.
    """
    info = etapas.get((pipeline_id, status_id))
    if info is None:
        info = etapas.get(status_id)
    if info is None and status_id is not None:
        info = next((v for k, v in etapas.items()
                     if _status_id(k) == status_id), None)
    return info or SIN_ETAPA


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


def resolver_embudos(etapas_cat, nombres_embudo):
    """Arma la config de cada embudo ya con su pipeline_id y sus etapas.

    Devuelve {pipeline_id: {...}} en el orden de configuration.json
    (el primero es el principal).
    """
    resuelto = {}
    for orden, emb in enumerate(EMBUDOS_CFG):
        nombre = emb["NOMBRE"]

        # El id sale de secret.json (PIPELINE_ID) o del propio config (ID)
        if emb.get("ID"):
            pid = int(emb["ID"])
        else:
            clave = emb.get("CLAVE_SECRET", nombre)
            try:
                pid = int(secret["kommo"]["PIPELINE_ID"][clave])
            except KeyError:
                sys.exit(f"ERROR: falta PIPELINE_ID['{clave}'] en secret.json")

        todas = a_bool(emb.get("TODAS_LAS_ETAPAS"), False)

        # Etapas que se aceptan de este embudo cuando NO se traen todas:
        #   ETAPAS_EXACTAS    -> el nombre debe ser igual (sin acentos/mayus)
        #   ETAPAS_CONTIENEN  -> el nombre debe contener ese texto
        exactas = {normalizar(x) for x in emb.get("ETAPAS_EXACTAS", [])}
        contienen = [normalizar(x) for x in emb.get("ETAPAS_CONTIENEN", [])]

        etapas_ids, etapas_nombres = set(), []
        if not todas:
            for clave, (p_id, _emb, etapa, _orden) in etapas_cat.items():
                if p_id != pid:
                    continue
                sid = _status_id(clave)
                n = normalizar(etapa)
                if n in exactas or any(c in n for c in contienen):
                    etapas_ids.add(sid)
                    etapas_nombres.append(etapa)

        # Campos obligatorios: los del embudo si los trae, si no los globales
        requeridos = emb.get("CAMPOS_REQUERIDOS", CAMPOS_REQUERIDOS)

        resuelto[pid] = {
            "nombre": nombre,
            "nombre_kommo": nombres_embudo.get(pid, nombre),
            "orden": orden,
            "principal": orden == 0,
            "todas": todas,
            "etapas_ids": etapas_ids,
            "etapas_nombres": sorted(etapas_nombres),
            "campos_requeridos": list(requeridos),
        }
        if todas:
            log(f"  {nombre} ({pid}): todas las etapas")
        else:
            log(f"  {nombre} ({pid}): {len(etapas_ids)} etapas -> "
                f"{', '.join(sorted(etapas_nombres)) or 'NINGUNA'}")
    return resuelto


def _valor_campo(cf):
    """Convierte un custom_fields_values de Kommo a un valor de Excel."""
    valores = []
    for v in (cf.get("values") or []):
        x = v.get("value")
        if x is None or x == "":
            continue
        tipo = cf.get("field_type")
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            # las fechas de Kommo viajan como timestamp
            if tipo in ("date", "birthday"):
                x = datetime.fromtimestamp(x, TZ).date()
            elif tipo == "date_time":
                x = datetime.fromtimestamp(x, TZ).replace(tzinfo=None)
        elif isinstance(x, bool):
            x = "Si" if x else "No"
        valores.append(x)
    if not valores:
        return None
    if len(valores) == 1:
        return valores[0]
    return " | ".join(str(v) for v in valores)     # campos de opcion multiple


def campos_pedidos():
    """Nombres normalizados de los campos que hay que leer de la tarjeta."""
    pedidos = set(CAMPOS_TARJETA)
    pedidos.update(CAMPOS_REQUERIDOS)
    for emb in EMBUDOS_CFG:
        pedidos.update(emb.get("CAMPOS_REQUERIDOS", []))
    return {normalizar(c) for c in pedidos}


def cargar_tarjetas(lead_ids):
    """Devuelve {lead_id: {"campos": {...}, "etiquetas": "a, b"}}.

    Se piden en lotes de 250 con filter[id][], no uno por uno.
    """
    quiero = campos_pedidos()
    tarjetas = {}
    if (not quiero and not INCLUIR_ETIQUETAS) or not lead_ids:
        return tarjetas

    ids = sorted(lead_ids)
    LOTE = 250
    for i in range(0, len(ids), LOTE):
        lote = ids[i:i + LOTE]
        page = 1
        while True:
            data = get(f"{BASE}/leads",
                       {"filter[id][]": lote, "limit": LOTE, "page": page})
            if not data:
                break
            for lead in data["_embedded"]["leads"]:
                datos = {}
                for cf in (lead.get("custom_fields_values") or []):
                    nombre = normalizar(cf.get("field_name") or "")
                    if nombre in quiero:
                        datos[nombre] = _valor_campo(cf)
                etiquetas = [t.get("name") for t in
                             (lead.get("_embedded") or {}).get("tags") or []
                             if t.get("name")]
                tarjetas[lead["id"]] = {"campos": datos,
                                        "etiquetas": ", ".join(etiquetas)}
            if not data.get("_links", {}).get("next"):
                break
            page += 1
        log(f"  tarjetas leidas: {len(tarjetas)}/{len(ids)}")
    return tarjetas


def aplicar_campos(filas, tarjetas, embudos):
    """Agrega las columnas de la tarjeta y descarta los leads incompletos."""
    salida, descartados = [], set()
    for fila in filas:
        datos = tarjetas.get(fila["LEAD_ID"], {}).get("campos", {})
        requeridos = embudos[fila["PIPELINE_ID"]]["campos_requeridos"]

        faltante = any(datos.get(normalizar(c)) in (None, "")
                       for c in requeridos)
        if faltante:
            descartados.add(fila["LEAD_ID"])
            continue

        for etiqueta in CAMPOS_TARJETA:
            fila[etiqueta] = datos.get(normalizar(etiqueta))
        salida.append(fila)
    return salida, len(descartados)


def marcar_etiquetas(df, etiquetas):
    """Escribe las etiquetas SOLO en la fila mas reciente de cada lead.

    Se hace por archivo, asi que si un lead sale en dos archivos, en cada
    uno queda marcada su propia fila mas reciente.
    """
    df = df.copy()
    df[COL_ETIQUETAS] = None
    if df.empty:
        return df
    ultimas = df.groupby("LEAD_ID")["FECHA"].idxmax()
    df.loc[ultimas, COL_ETIQUETAS] = [
        etiquetas.get(lead_id) or None
        for lead_id in df.loc[ultimas, "LEAD_ID"]]
    return df


def leads_en_etapas(pipeline_id, status_ids):
    """IDs de los leads que HOY estan parados en esas etapas."""
    ids = set()
    if not status_ids:
        return ids
    page = 1
    while True:
        params = {"page": page, "limit": 250}
        for i, sid in enumerate(sorted(status_ids)):
            params[f"filter[statuses][{i}][pipeline_id]"] = pipeline_id
            params[f"filter[statuses][{i}][status_id]"] = sid
        data = get(f"{BASE}/leads", params)
        if not data:
            break
        for lead in data["_embedded"]["leads"]:
            ids.add(lead["id"])
        if not data.get("_links", {}).get("next"):
            break
        page += 1
    return ids


# ----------------------------------------------------------------------
# EVENTOS
# ----------------------------------------------------------------------
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
    desde = a_timestamp(FECHA_DESDE)
    hasta = a_timestamp(FECHA_HASTA, fin_de_dia=True)

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
# ARMADO DEL DATAFRAME
# ----------------------------------------------------------------------
def armar_filas(eventos, etapas, usuarios, embudos, leads_permitidos):
    filas = []
    for ev in eventos:
        sid_ant, pid_ant = extraer_status(ev.get("value_before"))
        sid_new, pid_new = extraer_status(ev.get("value_after"))
        pipeline_id = pid_new or pid_ant

        cfg = embudos.get(pipeline_id)
        if cfg is None:          # embudo que no nos interesa
            continue

        lead_id = ev["entity_id"]

        if not cfg["todas"]:
            if FILTRO_POR_ETAPA_ACTUAL:
                # solo leads que hoy estan en las etapas pedidas
                if lead_id not in leads_permitidos.get(pipeline_id, set()):
                    continue
            else:
                # solo los movimientos hacia las etapas pedidas
                if sid_new not in cfg["etapas_ids"]:
                    continue

        _, _, eta_new, orden = buscar_etapa(etapas, pid_new, sid_new)
        _, _, eta_ant, _ = buscar_etapa(etapas, pid_ant, sid_ant)

        fecha = datetime.fromtimestamp(ev["created_at"], TZ).replace(
            tzinfo=None)

        fila = {
            "LEAD_ID": lead_id,
            "EMBUDO": cfg["nombre_kommo"],
            "PIPELINE_ID": pipeline_id,   # auxiliar: no se exporta
            "ETAPA_ANTERIOR": eta_ant,
            "ETAPA_NUEVA": eta_new,
            "FECHA": fecha,               # fecha y hora en una sola celda
            "TIPO_EVENTO": ev["type"],    # auxiliar: no se exporta
            "ORDEN_ETAPA": orden,
            "ORDEN_EMBUDO": cfg["orden"],
        }
        if INCLUIR_USUARIOS:
            fila["MOVIDO_POR"] = usuarios.get(ev.get("created_by"),
                                              ev.get("created_by"))
        filas.append(fila)
    return filas


def cols_historial():
    """Orden de las columnas de la hoja HISTORIAL.

    Los campos de la tarjeta (CAMPOS_TARJETA) van despues de EMBUDO,
    porque son datos del lead y no del movimiento.
    """
    return (["LEAD_ID", "EMBUDO"] + list(CAMPOS_TARJETA) +
            ([COL_ETIQUETAS] if INCLUIR_ETIQUETAS else []) +
            ["ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA",
             "MOVIDO_POR", "DIAS_EN_ETAPA_ANTERIOR"])


def construir_df(filas):
    df = pd.DataFrame(filas)

    # Dias entre un movimiento y el siguiente del mismo lead.
    # Se calcula SIEMPRE sobre el historial completo y en orden cronologico,
    # aunque el lead haya cruzado de un embudo a otro.
    df = df.sort_values(["LEAD_ID", "FECHA"])
    df["DIAS_EN_ETAPA_ANTERIOR"] = (
        df.groupby("LEAD_ID")["FECHA"].diff().dt.total_seconds() / 86400
    ).round(2)

    # Orden de salida: primero el embudo principal, luego los demas
    return df.sort_values(["ORDEN_EMBUDO", "LEAD_ID", "FECHA"])


def construir_pivote(df):
    """Una fila por lead, una columna por etapa con la primera fecha.

    Si una misma etapa existe en dos embudos (p. ej. "Logrado con exito"),
    la columna se etiqueta "EMBUDO - ETAPA" para no mezclarlas.
    """
    cols = (df[["EMBUDO", "ETAPA_NUEVA", "ORDEN_EMBUDO", "ORDEN_ETAPA"]]
            .drop_duplicates()
            .sort_values(["ORDEN_EMBUDO", "ORDEN_ETAPA", "ETAPA_NUEVA"]))

    repetidas = (cols.groupby("ETAPA_NUEVA")["EMBUDO"].nunique()
                 .loc[lambda s: s > 1].index)

    def etiqueta(embudo, etapa):
        return f"{embudo} - {etapa}" if etapa in repetidas else etapa

    df = df.copy()
    df["ETAPA_COL"] = [etiqueta(e, t)
                       for e, t in zip(df["EMBUDO"], df["ETAPA_NUEVA"])]

    orden_cols, vistas = [], set()
    for _, r in cols.iterrows():
        col = etiqueta(r["EMBUDO"], r["ETAPA_NUEVA"])
        if col not in vistas:
            vistas.add(col)
            orden_cols.append(col)

    return (df.pivot_table(index="LEAD_ID", columns="ETAPA_COL",
                           values="FECHA", aggfunc="min")
              .reindex(columns=orden_cols)
              .reset_index())


# ----------------------------------------------------------------------
# EXCEL
# ----------------------------------------------------------------------
def escribir_excel(ruta, df, etiquetas=None):
    os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
    if INCLUIR_ETIQUETAS:
        df = marcar_etiquetas(df, etiquetas or {})
    historial = df[[c for c in cols_historial() if c in df.columns]]
    pivote = construir_pivote(df)

    with pd.ExcelWriter(ruta, engine="openpyxl",
                        datetime_format="yyyy-mm-dd hh:mm:ss",
                        date_format="yyyy-mm-dd") as xl:
        historial.to_excel(xl, sheet_name="HISTORIAL", index=False)
        pivote.to_excel(xl, sheet_name="PIVOTE", index=False)

        for hoja, ancho in (("HISTORIAL", 22), ("PIVOTE", 20)):
            ws = xl.sheets[hoja]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = ancho
    return ruta


def nombre_por_embudo(ruta_base, nombre_embudo):
    raiz, ext = os.path.splitext(ruta_base)
    return f"{raiz}_{nombre_embudo}{ext}"


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    print(f"Kommo: {SUBDOMAIN} | Hilos: {HILOS} | Desde: {FECHA_DESDE} | "
          f"Hasta: {FECHA_HASTA or 'hoy'} | "
          f"Embudo principal: {EMBUDOS_CFG[0]['NOMBRE']} | "
          f"Embudo complementario: {EMBUDOS_CFG[1]['NOMBRE'] if len(EMBUDOS_CFG) > 1 else 'NINGUNO'} | "
          f"Archivos separados: {ARCHIVOS_SEPARADOS}")
    if not TOKEN:
        sys.exit("ERROR: falta el token de Kommo (variable de entorno "
                 "KOMMO_TOKEN o secret.json -> kommo -> TOKEN)")

    inicio = time.time()

    log("Cargando catalogo de embudos y etapas...")
    etapas, nombres_embudo = cargar_catalogo_etapas()
    embudos = resolver_embudos(etapas, nombres_embudo)
    usuarios = cargar_usuarios()

    # Leads que hoy estan en las etapas pedidas (solo embudos filtrados)
    leads_permitidos = {}
    if FILTRO_POR_ETAPA_ACTUAL:
        for pid, cfg in embudos.items():
            if not cfg["todas"]:
                leads_permitidos[pid] = leads_en_etapas(pid, cfg["etapas_ids"])
                log(f"  {cfg['nombre']}: {len(leads_permitidos[pid])} leads "
                    f"en las etapas pedidas")

    log("Descargando eventos...")
    eventos = descargar_eventos()
    if not eventos:
        sys.exit("No se encontraron eventos en el rango solicitado.")

    filas = armar_filas(eventos, etapas, usuarios, embudos, leads_permitidos)
    if not filas:
        sys.exit("No hubo movimientos que cumplan los filtros configurados.")

    # Datos de la tarjeta del lead (campos personalizados y etiquetas) +
    # descarte de los que tienen vacio alguno de los CAMPOS_REQUERIDOS
    etiquetas = {}
    if campos_pedidos() or INCLUIR_ETIQUETAS:
        log("Leyendo tarjetas de los leads...")
        tarjetas = cargar_tarjetas({f["LEAD_ID"] for f in filas})
        etiquetas = {lid: t["etiquetas"] for lid, t in tarjetas.items()}
        filas, descartados = aplicar_campos(filas, tarjetas, embudos)
        if descartados:
            log(f"  {descartados} leads descartados por campos vacios")
        if not filas:
            sys.exit("Ningun lead tiene llenos todos los CAMPOS_REQUERIDOS "
                     f"({', '.join(CAMPOS_REQUERIDOS) or 'sin campos'}). "
                     "Revisa que los nombres coincidan con los de Kommo.")

    df = construir_df(filas)

    # ---- Salida ------------------------------------------------------
    generados = []
    if ARCHIVOS_SEPARADOS:
        # Un archivo por embudo, en el orden de configuration.json
        for pid, cfg in sorted(embudos.items(), key=lambda x: x[1]["orden"]):
            parte = df[df["ORDEN_EMBUDO"] == cfg["orden"]]
            if parte.empty:
                continue
            ruta = nombre_por_embudo(SALIDA, cfg["nombre"])
            escribir_excel(ruta, parte, etiquetas)
            generados.append((ruta, parte))
    else:
        # Todo junto: primero el embudo principal, luego los demas
        escribir_excel(SALIDA, df, etiquetas)
        generados.append((SALIDA, df))

    limpiar_avance()
    for ruta, parte in generados:
        print(f"{fn.ruta_para_mostrar(ruta)} | {len(parte)} movimientos | "
              f"{parte['LEAD_ID'].nunique()} leads")
    resumen = " | ".join(f"{c['nombre_kommo']}: "
                         f"{(df['ORDEN_EMBUDO'] == c['orden']).sum()}"
                         for c in sorted(embudos.values(),
                                         key=lambda x: x["orden"]))
    print(f"{resumen} | {time.time() - inicio:.1f}s")


if __name__ == "__main__":
    main()