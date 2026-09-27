"""
Historial de etapas de Kommo (leads de los asesores de cierres)
===============================================================

Descarga de Kommo el historial completo de etapas de los leads que HOY tienen
en sus etiquetas a alguno de los asesores de secret.json -> asesores, y genera
un Excel con:

  - Hoja "HISTORIAL": un renglón por movimiento (la creación del lead y cada
    cambio de etapa o de embudo), con las columnas
        LEAD, EMBUDO, CREACION, ESTATUS DE NEGOCIO, FECHA DICTAMEN,
        MONTO OTORGADO, MOTIVO LEAD PERDIDO, ETIQUETAS, ETAPA ANTERIOR,
        ETAPA NUEVA, FECHA EVENTO, DIAS ETAPA ANTERIOR
  - Hoja "PIVOTE": un renglón por lead, una columna por "EMBUDO - ETAPA" con la
    primera fecha en que el lead entró a esa etapa, más EMBUDO ACTUAL y
    CAMBIOS DE EMBUDO.

Qué se descarga (sin límite de fechas):
  1. Leads de los embudos de secret.json -> kommo -> PIPELINE_ID (cualquier
     etapa) que tienen al menos una etiqueta que contiene el nombre de un
     asesor de secret.json -> asesores (sin importar mayúsculas ni acentos).
  2. Todos sus cambios de etapa dentro de esos embudos, y sus entradas y
     salidas hacia otros embudos (los movimientos enteramente en otros
     embudos no se guardan). La creación del lead es el primer renglón.
  3. Los campos de su tarjeta de secret.json -> campos_tarjeta (vacíos si no
     tienen dato). Los de campos_dinero van con formato de dinero.

Además genera, junto al historial, el reporte de leads anómalos:
  - leads_anomalos.txt  (para leerlo) y leads_anomalos.json (lo usan los
    reportes para asignar los leads con más de un asesor):
      * más de un asesor en sus etiquetas, con la fecha de cada etiqueta y el
        asesor asignado (el de la etiqueta más reciente; en empate, ninguno:
        el lead cuenta para todos),
      * con FECHA DICTAMEN pero sin ESTATUS DE NEGOCIO

Uso:
    python extract_data_from_kommo.py      # guarda en CARPETA_HISTORIAL

    Desde app.py se usa descargar_historial(), que informa el avance.
    El token se toma de la variable de entorno KOMMO_TOKEN si existe; si no,
    de secret.json -> kommo -> TOKEN.

configuration.json -> settings (solo configuración y rendimiento):
    CARPETA_HISTORIAL  carpeta donde se guarda el historial.
    performance        THREADS (descargas en paralelo), VERBOSE (detalle).
"""
import json
import os
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from pathlib import Path
from threading import Lock

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

import functions as fn

config = fn.import_json("json/configuration.json")
secret = fn.import_json("json/secret.json")



# ----------------------------------------------------------------------
# UTILIDADES
# ----------------------------------------------------------------------
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


# ----------------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------------
SUBDOMAIN = secret["kommo"]["SUBDOMAIN"]
TOKEN = os.environ.get("KOMMO_TOKEN") or secret["kommo"].get("TOKEN")

# Embudos: {clave: pipeline_id} de secret.json (la clave es el nombre que se
# muestra en la columna EMBUDO). El orden es el de secret.json.
EMBUDOS = {clave: int(pid) for clave, pid in secret["kommo"]["PIPELINE_ID"].items()}
ASESORES = list(secret.get("asesores") or [])
# Campos de la tarjeta (secret.json). Si faltan, los de la hoja HISTORIAL.
CAMPOS_TARJETA = list(secret.get("campos_tarjeta") or
                      ["ESTATUS DE NEGOCIO", "FECHA DICTAMEN", "MONTO OTORGADO", "MOTIVO LEAD PERDIDO"])
CAMPOS_DINERO = list(secret.get("campos_dinero") or ["MONTO OTORGADO"])
FORMATO_DINERO = '"$"#,##0.00'

_perf = config["settings"].get("performance", {})
VERBOSE = a_bool(_perf.get("VERBOSE"), False)
HILOS = max(1, int(_perf.get("THREADS", 4)))

TZ = timezone(timedelta(hours=-6))
BASE = f"https://{SUBDOMAIN}.kommo.com/api/v4"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}

COL_ETIQUETAS = "ETIQUETAS"
COL_CREACION = "CREACION"
# Encabezados de la hoja HISTORIAL (nombre interno -> título en el Excel).
# Los campos de la tarjeta usan tal cual su nombre de campos_tarjeta.
ENCABEZADOS_HISTORIAL = {
    "LEAD_ID": "LEAD",
    "ETAPA_ANTERIOR": "ETAPA ANTERIOR",
    "ETAPA_NUEVA": "ETAPA NUEVA",
    "FECHA": "FECHA EVENTO",
    "DIAS_EN_ETAPA_ANTERIOR": "DIAS ETAPA ANTERIOR",
}

# Una sola sesión = una sola negociación TLS reutilizada en todas las llamadas.
# Los reintentos se hacen solo en get().
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
SESSION.mount("https://", HTTPAdapter(pool_connections=HILOS, pool_maxsize=HILOS * 2))
SESSION.headers.update({"User-Agent": "kommo-export/1.0"})

# Aviso de avance para interfaces gráficas (app.py): función(fraccion, texto)
# con fraccion entre 0 y 1 (None = solo actualizar el texto).
AL_AVANZAR = None
_lock = Lock()



def log(msg):
    if VERBOSE:
        print(msg, file=sys.stderr)


def _avisar(fraccion, texto):
    if AL_AVANZAR is not None:
        try:
            AL_AVANZAR(fraccion, texto)
        except Exception:  # noqa: BLE001 - un fallo de la interfaz no debe detener la descarga
            pass


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
# CATÁLOGO DE EMBUDOS Y ETAPAS
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


def extraer_status(bloque):
    """Saca (status_id, pipeline_id) de value_before / value_after."""
    if not bloque:
        return None, None
    item = bloque[0] if isinstance(bloque, list) else bloque
    st = item.get("lead_status") or {}
    return st.get("id"), st.get("pipeline_id")


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


def a_numero(valor):
    """'12,345.50' / '$12345.5' / 12345.5 -> 12345.5 (vacío -> None).

    Si no se puede leer como número, se deja el texto tal cual.
    """
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return float(valor)
    texto = str(valor).replace("$", "").replace(",", "").replace(" ", "").strip()
    try:
        return float(texto)
    except ValueError:
        return valor


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


def cambios_de_embudo(grupo):
    """'VENTAS → CIERRES (07/09/2026 10:00)' por cada cambio de embudo del lead."""
    cambios = []
    for ant, emb, fecha in zip(grupo["EMBUDO_ANTERIOR"], grupo["EMBUDO"], grupo["FECHA"]):
        if ant and isinstance(ant, str) and ant != emb:
            cambios.append(f"{ant} → {emb} ({fecha:%d/%m/%Y %H:%M})")
    return "; ".join(cambios) or None


def construir_pivote(df):
    """Una fila por lead, una columna por etapa con la primera fecha.

    Cada columna de etapa se titula "EMBUDO - ETAPA", para que se vea en qué
    embudo estuvo el lead (p. ej. "VENTAS - Contactado", "CIERRES - Oferta").
    Después de LEAD_ID van EMBUDO ACTUAL (el del último movimiento) y CAMBIOS
    DE EMBUDO (cada cambio con su fecha).
    """
    cols = (df[["EMBUDO", "ETAPA_NUEVA", "ORDEN_EMBUDO", "ORDEN_ETAPA"]]
            .drop_duplicates()
            .sort_values(["ORDEN_EMBUDO", "ORDEN_ETAPA", "ETAPA_NUEVA"]))

    def etiqueta(embudo, etapa):
        return f"{embudo} - {etapa}"

    df = df.copy()
    df["ETAPA_COL"] = [etiqueta(e, t)
                       for e, t in zip(df["EMBUDO"], df["ETAPA_NUEVA"])]

    orden_cols, vistas = [], set()
    for _, r in cols.iterrows():
        col = etiqueta(r["EMBUDO"], r["ETAPA_NUEVA"])
        if col not in vistas:
            vistas.add(col)
            orden_cols.append(col)

    pivote = (df.pivot_table(index="LEAD_ID", columns="ETAPA_COL",
                             values="FECHA", aggfunc="min")
                .reindex(columns=orden_cols))

    ordenado = df.sort_values(["LEAD_ID", "FECHA"])
    if "EMBUDO_ANTERIOR" not in ordenado.columns:
        ordenado["EMBUDO_ANTERIOR"] = None
    por_lead = ordenado.groupby("LEAD_ID")
    pivote.insert(0, "EMBUDO ACTUAL", por_lead["EMBUDO"].last())
    pivote.insert(1, "CAMBIOS DE EMBUDO",
                  por_lead[["EMBUDO_ANTERIOR", "EMBUDO", "FECHA"]].apply(cambios_de_embudo))
    pivote.columns.name = None
    return pivote.reset_index()


# ----------------------------------------------------------------------
# EMBUDOS, ASESORES Y TARJETAS
# ----------------------------------------------------------------------
def resolver_embudos(nombres_embudo):
    """{pipeline_id: {"nombre": clave de secret.json, "orden": n}}.

    Un ID que no existe en Kommo descartaría en silencio todo el embudo, así
    que se detiene con un mensaje claro.
    """
    embudos = {}
    for orden, (clave, pid) in enumerate(EMBUDOS.items()):
        if pid not in nombres_embudo:
            disponibles = "; ".join(f"{n} ({i})" for i, n in nombres_embudo.items())
            sys.exit(f"ERROR: el embudo {clave} (ID {pid}) no existe en Kommo. "
                     f"Revisa secret.json -> kommo -> PIPELINE_ID. "
                     f"Embudos disponibles: {disponibles}")
        embudos[pid] = {"nombre": clave, "orden": orden}
    return embudos


def nombre_embudo(pid, embudos, nombres_embudo):
    """Clave de secret.json para los embudos configurados; nombre en Kommo para los demás."""
    if pid is None:
        return None
    if pid in embudos:
        return embudos[pid]["nombre"]
    return nombres_embudo.get(pid, str(pid))


def asesores_en_etiquetas(etiquetas):
    """Asesores de secret.json cuyo nombre está contenido en alguna etiqueta
    (sin importar mayúsculas ni acentos). Si en una misma etiqueta caben dos
    asesores porque un nombre contiene al otro ("ANA X" dentro de "JULIANA X"),
    cuenta solo el más largo."""
    encontrados = set()
    for etiqueta in etiquetas:
        texto = normalizar(etiqueta)
        caben = [a for a in ASESORES if normalizar(a) in texto]
        encontrados.update(a for a in caben
                           if not any(b != a and normalizar(a) in normalizar(b) for b in caben))
    return [a for a in ASESORES if a in encontrados]


def _fecha(ts):
    """Timestamp de Kommo -> fecha y hora local (sin zona)."""
    return datetime.fromtimestamp(ts, TZ).replace(tzinfo=None) if ts else None


def _paginar(url, params, clave):
    """Todas las páginas de un listado de Kommo."""
    params, items = dict(params), []
    params["page"] = 1
    while True:
        data = get(url, params)
        if not data:
            break
        items.extend(data["_embedded"][clave])
        if not data.get("_links", {}).get("next"):
            break
        params["page"] += 1
    return items


def _tarjeta(lead):
    """Datos que se guardan de la tarjeta de un lead."""
    quiero = {normalizar(c) for c in CAMPOS_TARJETA}
    datos = {}
    for cf in lead.get("custom_fields_values") or []:
        nombre = normalizar(cf.get("field_name"))
        if nombre in quiero:
            datos[nombre] = _valor_campo(cf)
    campos = {}
    for campo in CAMPOS_TARJETA:
        valor = datos.get(normalizar(campo))
        campos[campo] = a_numero(valor) if campo in CAMPOS_DINERO else valor
    etiquetas = [t.get("name") for t in ((lead.get("_embedded") or {}).get("tags") or [])
                 if t.get("name")]
    return {"campos": campos, "etiquetas": etiquetas,
            "asesores": asesores_en_etiquetas(etiquetas),
            "creacion": _fecha(lead.get("created_at")),
            "pipeline_id": lead.get("pipeline_id"), "status_id": lead.get("status_id")}


def identificar_leads(etapas, embudos):
    """Paso 1: leads de los embudos (en cualquier etapa, sin límite de fechas)
    con al menos una etiqueta de un asesor. Devuelve {LEAD_ID: tarjeta}."""
    pares = sorted(clave for clave in etapas
                   if isinstance(clave, tuple) and clave[0] in embudos)
    tandas = [pares[i:i + 10] for i in range(0, len(pares), 10)]

    def bajar(tanda):
        params = {"limit": 250}
        for i, (pid, sid) in enumerate(tanda):
            params[f"filter[statuses][{i}][pipeline_id]"] = pid
            params[f"filter[statuses][{i}][status_id]"] = sid
        return _paginar(f"{BASE}/leads", params, "leads")

    leads = {}
    with ThreadPoolExecutor(max_workers=HILOS) as ex:
        for hechos, lote in enumerate(ex.map(bajar, tandas), start=1):
            for lead in lote:
                leads[lead["id"]] = lead
            _avisar(0.03 + 0.22 * hechos / len(tandas),
                    f"Buscando leads de los asesores... {len(leads)} revisados")
    tarjetas = {lid: _tarjeta(lead) for lid, lead in leads.items()}
    log(f"  {len(leads)} leads en los embudos")
    return {lid: t for lid, t in tarjetas.items() if t["asesores"]}


def eventos_de_leads(lead_ids, tipos, avance=None):
    """Eventos de esos leads (sin límite de fechas), de 10 en 10 leads."""
    ids = sorted(lead_ids)
    tandas = [ids[i:i + 10] for i in range(0, len(ids), 10)]

    def bajar(tanda):
        params = {"filter[entity][]": "lead", "filter[entity_id][]": tanda, "limit": 100}
        for i, tipo in enumerate(tipos):
            params[f"filter[type][{i}]"] = tipo
        return _paginar(f"{BASE}/events", params, "events")

    eventos = {}
    with ThreadPoolExecutor(max_workers=HILOS) as ex:
        for hechos, lote in enumerate(ex.map(bajar, tandas), start=1):
            for ev in lote:
                eventos[ev["id"]] = ev
            if avance:
                inicio, fin = avance
                _avisar(inicio + (fin - inicio) * hechos / len(tandas),
                        f"Descargando movimientos... {hechos * 10 if hechos < len(tandas) else len(ids)}"
                        f"/{len(ids)} leads")
    return list(eventos.values())


# ----------------------------------------------------------------------
# RENGLONES DEL HISTORIAL
# ----------------------------------------------------------------------
def _fila(lead_id, pid_ant, sid_ant, pid_new, sid_new, fecha, tipo, orden_evento,
          etapas, embudos, nombres_embudo):
    _, _, eta_new, orden = buscar_etapa(etapas, pid_new, sid_new)
    eta_ant = buscar_etapa(etapas, pid_ant, sid_ant)[2] if sid_ant is not None else None
    return {
        "LEAD_ID": lead_id,
        "EMBUDO": nombre_embudo(pid_new, embudos, nombres_embudo),
        "ETAPA_ANTERIOR": eta_ant or None,
        "ETAPA_NUEVA": eta_new,
        "FECHA": fecha,
        # auxiliares (no se exportan al HISTORIAL)
        "EMBUDO_ANTERIOR": nombre_embudo(pid_ant, embudos, nombres_embudo),
        "ORDEN_EMBUDO": embudos.get(pid_new, {}).get("orden", len(embudos)),
        "ORDEN_ETAPA": orden,
        "ORDEN_EVENTO": orden_evento,
        "TIPO_EVENTO": tipo,
    }


def armar_filas(eventos, tarjetas, etapas, embudos, nombres_embudo):
    """Creación del lead + cambios de etapa/embudo que tocan los embudos
    configurados (dentro de ellos, o entrando/saliendo de ellos)."""
    por_lead = {}
    for ev in eventos:
        if ev.get("entity_id") in tarjetas:
            por_lead.setdefault(ev["entity_id"], []).append(ev)

    filas = []
    for lid, tarjeta in tarjetas.items():
        evs = sorted(por_lead.get(lid, []), key=lambda e: e["created_at"])
        cambios = [e for e in evs if e.get("type") == "lead_status_changed"]
        alta = next((e for e in evs if e.get("type") == "lead_added"), None)

        # Creación: etapa inicial = la del alta; si Kommo no tiene el evento,
        # la etapa previa al primer cambio; si no hubo cambios, la actual.
        if alta:
            sid0, pid0 = extraer_status(alta.get("value_after"))
        elif cambios:
            sid0, pid0 = extraer_status(cambios[0].get("value_before"))
        else:
            sid0, pid0 = tarjeta["status_id"], tarjeta["pipeline_id"]
        creacion = tarjeta["creacion"] or (_fecha(alta["created_at"]) if alta else None)
        if creacion and pid0 in embudos:
            filas.append(_fila(lid, None, None, pid0, sid0, creacion, "lead_added", 0,
                               etapas, embudos, nombres_embudo))

        for n, ev in enumerate(cambios, start=1):
            sid_ant, pid_ant = extraer_status(ev.get("value_before"))
            sid_new, pid_new = extraer_status(ev.get("value_after"))
            if pid_new not in embudos and pid_ant not in embudos:
                continue      # movimiento enteramente en otros embudos
            filas.append(_fila(lid, pid_ant, sid_ant, pid_new, sid_new, _fecha(ev["created_at"]),
                               "lead_status_changed", n, etapas, embudos, nombres_embudo))
    return filas


def aplicar_tarjetas(filas, tarjetas):
    """Agrega a cada renglón la fecha de creación y los campos de la tarjeta."""
    for fila in filas:
        tarjeta = tarjetas[fila["LEAD_ID"]]
        fila[COL_CREACION] = tarjeta["creacion"]
        fila.update(tarjeta["campos"])
    return filas


def construir_df(filas):
    """DataFrame en orden cronológico por lead, con los días en la etapa anterior."""
    df = pd.DataFrame(filas).sort_values(["LEAD_ID", "FECHA", "ORDEN_EVENTO"], kind="mergesort")
    df["DIAS_EN_ETAPA_ANTERIOR"] = (
        df.groupby("LEAD_ID")["FECHA"].diff().dt.total_seconds() / 86400).round(2)
    return df.reset_index(drop=True)


# ----------------------------------------------------------------------
# LEADS ANÓMALOS
# ----------------------------------------------------------------------
CAMPO_DICTAMEN = "FECHA DICTAMEN"
CAMPO_ESTATUS = "ESTATUS DE NEGOCIO"
ARCHIVO_ANOMALIAS_TXT = "leads_anomalos.txt"
ARCHIVO_ANOMALIAS_JSON = "leads_anomalos.json"
# Resultado de la última descarga (lo usa app.py):
#   {"varias_asesores": {LEAD_ID: {"asesores": {asesor: fecha|None}, "asignado": asesor|None}},
#    "dictamen_sin_estatus": [LEAD_IDs]}
ULTIMAS_ANOMALIAS = {}


def _vacio(valor):
    return valor is None or (isinstance(valor, str) and not valor.strip())


def revisar_anomalias(tarjetas):
    """Clasifica los leads anómalos (sin las fechas de etiquetas)."""
    varias, con_dictamen = {}, []
    for lid in sorted(tarjetas):
        t = tarjetas[lid]
        if len(t["asesores"]) > 1:
            varias[lid] = {"asesores": {a: None for a in t["asesores"]}, "asignado": None}
        sin_dictamen = _vacio(t["campos"].get(CAMPO_DICTAMEN))
        sin_estatus = _vacio(t["campos"].get(CAMPO_ESTATUS))
        if sin_estatus and not sin_dictamen:
            con_dictamen.append(lid)
    return {"varias_asesores": varias, "dictamen_sin_estatus": con_dictamen}


def fechar_asesores(varias, tarjetas):
    """Fecha en que se puso la etiqueta de cada asesor (eventos 'etiqueta
    agregada', sin límite de fechas) y asesor asignado: el de la etiqueta más
    reciente. En empate, o si no hay fechas, no se asigna (cuenta para todos).
    Una etiqueta sin evento se considera más antigua que las que sí lo tienen.
    """
    ultima = {}                                    # (LEAD_ID, asesor) -> timestamp
    for ev in eventos_de_leads(varias, ["entity_tag_added"]):
        for item in ev.get("value_after") or []:
            nombre = ((item or {}).get("tag") or {}).get("name")
            for asesor in asesores_en_etiquetas([nombre] if nombre else []):
                clave = (ev["entity_id"], asesor)
                ultima[clave] = max(ultima.get(clave, 0), ev["created_at"])
    for lid, datos in varias.items():
        fechas = {a: ultima.get((lid, a)) for a in tarjetas[lid]["asesores"]}
        conocidas = {a: f for a, f in fechas.items() if f}
        mas_reciente = max(conocidas.values(), default=None)
        ganadores = [a for a, f in conocidas.items() if f == mas_reciente]
        datos["asesores"] = {a: _fecha(f) for a, f in fechas.items()}
        datos["asignado"] = ganadores[0] if len(ganadores) == 1 else None
    return varias


def texto_anomalias(anomalias):
    """Reporte de leads anómalos para leer (consola, ventana y .txt)."""
    varias = anomalias.get("varias_asesores") or {}
    lineas = [f"Leads con más de un asesor de cierres en sus etiquetas ({len(varias)}):"]
    for lid, datos in varias.items():
        fechas = ", ".join(f"{a} ({f:%d/%m/%Y %H:%M})" if f else f"{a} (sin fecha)"
                           for a, f in datos["asesores"].items())
        destino = (f"asignado a {datos['asignado']}" if datos["asignado"]
                   else "empate: cuenta para todos")
        lineas.append(f"  {lid}: {fechas} -> {destino}")
    for clave, titulo in (("dictamen_sin_estatus", "Leads con FECHA DICTAMEN pero sin ESTATUS DE NEGOCIO"),):
        ids = anomalias.get(clave) or []
        lineas.append(f"\n{titulo} ({len(ids)}):")
        if ids:
            lineas.append("  " + ", ".join(map(str, ids)))
    return "\n".join(lineas)


def guardar_anomalias(carpeta, anomalias):
    """Escribe leads_anomalos.txt (para leer) y leads_anomalos.json (para los reportes)."""
    carpeta = Path(carpeta)
    ahora = datetime.now()
    txt = f"Descarga del {ahora:%d/%m/%Y %H:%M}\n\n{texto_anomalias(anomalias)}\n"
    datos = {
        "generado": ahora.isoformat(timespec="seconds"),
        "varias_asesores": [
            {"lead": lid,
             "asesores": [{"asesor": a, "fecha_etiqueta": f.isoformat() if f else None}
                          for a, f in d["asesores"].items()],
             "asignado": d["asignado"]}
            for lid, d in (anomalias.get("varias_asesores") or {}).items()],
        "dictamen_sin_estatus": list(anomalias.get("dictamen_sin_estatus") or []),
    }
    try:
        # utf-8-sig: el Bloc de notas de Windows muestra bien los acentos
        (carpeta / ARCHIVO_ANOMALIAS_TXT).write_text(txt, encoding="utf-8-sig")
        (carpeta / ARCHIVO_ANOMALIAS_JSON).write_text(
            json.dumps(datos, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError as e:
        print(f"Aviso: no se pudieron guardar las anomalías en {carpeta}: {e}")
        return None
    print(f"Anomalías guardadas en: {carpeta / ARCHIVO_ANOMALIAS_TXT}")
    return carpeta / ARCHIVO_ANOMALIAS_TXT


# ----------------------------------------------------------------------
# EXCEL
# ----------------------------------------------------------------------
def cols_historial():
    """Columnas de la hoja HISTORIAL (nombres internos), en orden."""
    return (["LEAD_ID", "EMBUDO", COL_CREACION] + list(CAMPOS_TARJETA) +
            [COL_ETIQUETAS, "ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA", "DIAS_EN_ETAPA_ANTERIOR"])


def escribir_excel(ruta, df, etiquetas):
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    df = marcar_etiquetas(df, etiquetas)
    for col in cols_historial():
        if col not in df.columns:
            df[col] = None
    historial = df[cols_historial()].rename(columns=ENCABEZADOS_HISTORIAL)
    pivote = construir_pivote(df)

    with pd.ExcelWriter(ruta, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss",
                        date_format="yyyy-mm-dd") as xl:
        historial.to_excel(xl, sheet_name="HISTORIAL", index=False)
        pivote.to_excel(xl, sheet_name="PIVOTE", index=False)
        for hoja, ancho in (("HISTORIAL", 22), ("PIVOTE", 20)):
            ws = xl.sheets[hoja]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = ancho
        ws = xl.sheets["HISTORIAL"]
        for idx, celda in enumerate(ws[1], start=1):
            if celda.value in CAMPOS_DINERO:
                for (c,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                    if isinstance(c.value, (int, float)):
                        c.number_format = FORMATO_DINERO
    return ruta


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main(salida=None, guardar_anomalias_junto=True):
    """Descarga todo y escribe el historial en `salida` (por defecto, la ruta
    de configuration.json -> CARPETA_HISTORIAL)."""
    global ULTIMAS_ANOMALIAS
    ULTIMAS_ANOMALIAS = {}
    if not TOKEN:
        sys.exit("ERROR: falta el token de Kommo (variable de entorno "
                 "KOMMO_TOKEN o secret.json -> kommo -> TOKEN)")
    if not ASESORES:
        sys.exit("ERROR: secret.json -> asesores está vacío.")
    salida = Path(salida or fn.ruta_historial())
    inicio = time.time()
    print(f"Kommo: {SUBDOMAIN} | Embudos: {', '.join(EMBUDOS)} | "
          f"Asesores: {len(ASESORES)} | Hilos: {HILOS}")

    _avisar(0.02, "Conectando con Kommo...")
    etapas, nombres = cargar_catalogo_etapas()
    embudos = resolver_embudos(nombres)

    tarjetas = identificar_leads(etapas, embudos)
    if not tarjetas:
        sys.exit("Ningún lead de los embudos tiene en sus etiquetas a un asesor de secret.json.")
    print(f"Leads con etiqueta de un asesor: {len(tarjetas)}")

    _avisar(0.25, f"Descargando los movimientos de {len(tarjetas)} leads...")
    eventos = eventos_de_leads(tarjetas, ["lead_status_changed", "lead_added"], avance=(0.25, 0.85))
    filas = armar_filas(eventos, tarjetas, etapas, embudos, nombres)
    if not filas:
        sys.exit("No hubo movimientos de esos leads en los embudos configurados.")

    anomalias = revisar_anomalias(tarjetas)
    if anomalias["varias_asesores"]:
        _avisar(0.87, "Revisando leads con varios asesores...")
        fechar_asesores(anomalias["varias_asesores"], tarjetas)
    ULTIMAS_ANOMALIAS = anomalias

    _avisar(0.92, "Armando el historial...")
    df = construir_df(aplicar_tarjetas(filas, tarjetas))
    etiquetas = {lid: ", ".join(t["etiquetas"]) for lid, t in tarjetas.items()}
    _avisar(0.95, "Guardando el archivo...")
    escribir_excel(salida, df, etiquetas)

    print(f"{fn.ruta_para_mostrar(salida)} | {len(df)} movimientos | "
          f"{df['LEAD_ID'].nunique()} leads | {time.time() - inicio:.1f}s")
    print(texto_anomalias(anomalias))
    if guardar_anomalias_junto:
        guardar_anomalias(salida.parent, anomalias)
    _avisar(1.0, "Descarga terminada")
    return salida


def descargar_historial(salida=None, al_avanzar=None):
    """Descarga el historial para app.py y lo guarda en `salida` (por defecto,
    la ruta de configuration.json), junto con el reporte de leads anómalos.

    Se escribe primero con otro nombre y solo reemplaza al anterior si la
    descarga termina bien. Si falla, lanza RuntimeError con el motivo.
    """
    global AL_AVANZAR
    salida = Path(salida or fn.ruta_historial())
    salida.parent.mkdir(parents=True, exist_ok=True)
    temporal = salida.with_name(f"{salida.stem}.descargando{salida.suffix}")
    anterior, AL_AVANZAR = AL_AVANZAR, al_avanzar
    try:
        try:
            main(temporal, guardar_anomalias_junto=False)
        except SystemExit as e:        # main() termina con sys.exit("motivo")
            raise RuntimeError(str(e.code) if e.code else "La descarga se detuvo.") from None
        os.replace(temporal, salida)   # reemplaza por completo al anterior
        guardar_anomalias(salida.parent, ULTIMAS_ANOMALIAS)
        return salida
    finally:
        AL_AVANZAR = anterior
        if temporal.exists():
            try:
                temporal.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    main()
