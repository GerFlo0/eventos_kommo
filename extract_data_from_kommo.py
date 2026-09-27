#!/usr/bin/env python3
"""
Extrae el historial de cambios de etapa de los leads de Kommo (API v4)
y genera un Excel con:
  - Hoja "HISTORIAL": un renglon por cada cambio de etapa, con las columnas
        LEAD ID, EMBUDO, FECHA CREACION DE LEAD, ESTATUS DE NEGOCIO,
        FECHA DICTAMEN, MONTO OTORGADO, MOTIVO LEAD PERDIDO, ETIQUETAS,
        ETAPA ANTERIOR, ETAPA_NUEVA, FECHA EVENTO, DIAS EN ETAPA ANTERIOR
    (FECHA EVENTO lleva fecha y hora; MONTO OTORGADO va con formato de dinero)
  - Hoja "PIVOTE":    un renglon por lead, una columna por etapa
                      con la PRIMERA fecha en que el lead entro a esa etapa

Que se extrae:
  - Los cambios de etapa de CUALQUIER etapa de los embudos configurados
    (CIERRES y VENTAS) ...
  - ... de los leads que tienen al menos una etiqueta que contiene
    ETIQUETA_LEADS ("CIERRES"; sin importar mayusculas ni acentos).
  - Los campos de la tarjeta de CAMPOS_TARJETA; si estan vacios, la celda
    queda vacia (el lead no se descarta).

Al terminar se muestra en consola (y en la ventana de app.py) que leads con
etiqueta CIERRES tienen alguna anomalia, agrupados por tipo:
  - mas de una etiqueta con "CIERRES",
  - sin FECHA DICTAMEN,
  - sin ESTATUS DE NEGOCIO.

Requisitos:
    Ejecutar antes setup_invironment.py (instala requirements.txt).

Uso:
    python extract_data_from_kommo.py

    Desde app.py se usa descargar_historial(), que recibe el rango de fechas
    y la ruta de salida, e informa el avance para la barra de progreso.

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
Nota: la columna DIAS EN ETAPA ANTERIOR se calcula con TODOS los
movimientos del lead que quedan en el reporte (aunque haya cruzado de un
embudo a otro), asi que da el mismo numero en los dos modos. El primer
movimiento de cada lead dentro del rango de fechas queda vacio.
----------------------------------------------------------------------
"""

import os
import sys
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta
from pathlib import Path
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

# Solo entran los leads con al menos una etiqueta que CONTENGA este texto
# (sin importar mayusculas ni acentos). De ellos se traen sus cambios de
# etapa en cualquier etapa de los embudos configurados.
ETIQUETA_LEADS = config["settings"].get("ETIQUETA_LEADS", "CIERRES")

# ---- Campos de la tarjeta del lead ---------------------------------------
# CAMPOS_TARJETA    = campos personalizados que se agregan como columnas
#                     en la hoja HISTORIAL (se escriben tal cual aparecen
#                     en Kommo; no importan acentos ni mayusculas).
# CAMPOS_REQUERIDOS = de esos campos, los que el lead DEBE tener llenos
#                     para aparecer en el reporte. Si un lead tiene vacio
#                     alguno, se descarta completo.
#                     Lista vacia [] = no se filtra por campos (default).
# Se puede sobreescribir por embudo: basta con poner CAMPOS_REQUERIDOS
# dentro del bloque del embudo en configuration.json (por ejemplo [] en
# VENTAS para exigirlos solo en CIERRES).
CAMPOS_TARJETA = list(config["settings"].get("CAMPOS_TARJETA", []))
CAMPOS_REQUERIDOS = list(config["settings"].get("CAMPOS_REQUERIDOS", []))
# Campos de la tarjeta que son montos: se guardan como numero con formato de
# dinero en el Excel.
CAMPOS_DINERO = list(config["settings"].get("CAMPOS_DINERO", []))
FORMATO_DINERO = '"$"#,##0.00'

# Columna ETIQUETAS: las etiquetas (tags) del lead. Para no repetirlas en
# cada renglon, solo se escriben en la fila de la FECHA MAS RECIENTE de ese
# lead; si el lead tiene una sola fila, van en esa. Se apaga con false.
INCLUIR_ETIQUETAS = a_bool(config["settings"].get("INCLUIR_ETIQUETAS"), True)
COL_ETIQUETAS = "ETIQUETAS"

# Columna con la fecha (y hora) en que se creo el lead en Kommo. Sale del
# campo created_at de la tarjeta, asi que existe aunque la creacion haya
# sido antes de FECHA_DESDE.
COL_CREACION = "creacion_de_lead"

# Encabezados de la hoja HISTORIAL (nombre interno -> titulo en el Excel).
# Los campos de la tarjeta usan tal cual su nombre de CAMPOS_TARJETA.
ENCABEZADOS_HISTORIAL = {
    "LEAD_ID": "LEAD ID",
    COL_CREACION: "FECHA CREACION DE LEAD",
    "ETAPA_ANTERIOR": "ETAPA ANTERIOR",
    "ETAPA_NUEVA": "ETAPA_NUEVA",
    "FECHA": "FECHA EVENTO",
    "DIAS_EN_ETAPA_ANTERIOR": "DIAS EN ETAPA ANTERIOR",
}

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

# Aviso de avance para interfaces gráficas (app.py): función(fraccion, texto)
# con fraccion entre 0 y 1 (None = solo actualizar el texto). Por consola
# queda en None y no hace nada.
AL_AVANZAR = None


def _avisar(fraccion, texto):
    if AL_AVANZAR is not None:
        try:
            AL_AVANZAR(fraccion, texto)
        except Exception:  # noqa: BLE001 - un fallo de la interfaz no debe detener la descarga
            pass


def avance(n_eventos):
    with _lock:
        _prog["ev"] += n_eventos
        print(f"\r {_prog['ev']} eventos",
              end="", file=sys.stderr, flush=True)
        _avisar(None, f"Descargando eventos... {_prog['ev']} encontrados")


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
    """Arma la config de cada embudo con su pipeline_id.

    Devuelve {pipeline_id: {...}} en el orden de configuration.json
    (el primero es el principal). De cada embudo se toman TODAS sus etapas.
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

        # Un ID que no existe en Kommo descartaria en silencio todo el embudo
        if pid not in nombres_embudo:
            disponibles = "; ".join(f"{n} ({i})" for i, n in nombres_embudo.items())
            sys.exit(f"ERROR: el embudo {nombre} (ID {pid}) no existe en Kommo. "
                     f"Revisa el ID en configuration.json / secret.json. "
                     f"Embudos disponibles: {disponibles}")

        # Campos obligatorios: los del embudo si los trae, si no los globales
        requeridos = emb.get("CAMPOS_REQUERIDOS", CAMPOS_REQUERIDOS)

        resuelto[pid] = {
            "nombre": nombre,
            "nombre_kommo": nombres_embudo.get(pid, nombre),
            "orden": orden,
            "principal": orden == 0,
            "campos_requeridos": list(requeridos),
        }
        log(f"  {nombre} ({pid}): todas las etapas")
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


def _fecha_creacion(lead):
    """created_at de la tarjeta del lead como fecha y hora local (sin zona)."""
    ts = lead.get("created_at")
    if not ts:
        return None
    return datetime.fromtimestamp(ts, TZ).replace(tzinfo=None)


def cargar_tarjetas(lead_ids):
    """Devuelve {lead_id: {"campos": {...}, "etiquetas": "a, b"}}.

    Se piden en lotes de 250 con filter[id][], no uno por uno.
    """
    quiero = campos_pedidos()
    tarjetas = {}
    if not lead_ids:
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
                                        "etiquetas": ", ".join(etiquetas),
                                        "lista_etiquetas": etiquetas,
                                        "creacion": _fecha_creacion(lead),
                                        # embudo y etapa en que esta HOY
                                        "pipeline_id": lead.get("pipeline_id"),
                                        "status_id": lead.get("status_id")}
            if not data.get("_links", {}).get("next"):
                break
            page += 1
        log(f"  tarjetas leidas: {len(tarjetas)}/{len(ids)}")
        hechos = min(i + LOTE, len(ids))
        _avisar(0.75 + 0.18 * hechos / len(ids), f"Leyendo tarjetas de los leads... {hechos}/{len(ids)}")
    return tarjetas


def aplicar_campos(filas, tarjetas, embudos):
    """Agrega las columnas de la tarjeta y descarta los leads incompletos."""
    salida, descartados = [], set()
    for fila in filas:
        tarjeta = tarjetas.get(fila["LEAD_ID"], {})
        datos = tarjeta.get("campos", {})
        requeridos = embudos[fila["PIPELINE_ID"]]["campos_requeridos"]

        faltante = any(datos.get(normalizar(c)) in (None, "")
                       for c in requeridos)
        if faltante:
            descartados.add(fila["LEAD_ID"])
            continue

        fila[COL_CREACION] = tarjeta.get("creacion")

        for etiqueta in CAMPOS_TARJETA:
            valor = datos.get(normalizar(etiqueta))
            if etiqueta in CAMPOS_DINERO:
                valor = a_numero(valor)
            fila[etiqueta] = valor
        salida.append(fila)
    return salida, len(descartados)


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


# ----------------------------------------------------------------------
# ETAPA ACTUAL: completar el historial hasta donde esta hoy cada lead
# ----------------------------------------------------------------------
def _estado_final(filas):
    """{LEAD_ID: (pipeline_id, status_id)} del movimiento más reciente de cada lead."""
    ultimo = {}
    for f in filas:
        clave = (f["FECHA"], f.get("TIPO_EVENTO") != "lead_added")   # empate: el cambio de etapa gana
        if f["LEAD_ID"] not in ultimo or clave >= ultimo[f["LEAD_ID"]][0]:
            ultimo[f["LEAD_ID"]] = (clave, (f["PIPELINE_ID"], f.get("STATUS_ID")))
    return {lid: estado for lid, (_, estado) in ultimo.items()}


def leads_que_no_cuadran(filas, tarjetas):
    """Leads cuyo embudo/etapa actual en Kommo no es el del último movimiento."""
    distintos = set()
    for lid, estado in _estado_final(filas).items():
        tarjeta = tarjetas.get(lid) or {}
        actual = (tarjeta.get("pipeline_id"), tarjeta.get("status_id"))
        if actual[1] is not None and actual != estado:
            distintos.add(lid)
    return distintos


def completar_estado_actual(filas, tarjetas, etapas, usuarios, embudos):
    """Agrega los movimientos posteriores a FECHA_HASTA de los leads cuya etapa
    actual en Kommo no es la del último movimiento del periodo (p. ej. pasaron de
    VENTAS a CIERRES después de FECHA_HASTA), para que su historial termine en
    su etapa actual. Devuelve (filas, leads completados, leads que aún no cuadran).
    """
    pendientes = leads_que_no_cuadran(filas, tarjetas)
    hasta = a_timestamp(FECHA_HASTA, fin_de_dia=True)
    ahora = int(time.time())
    if not pendientes or hasta is None or hasta >= ahora:
        return filas, set(), pendientes

    log(f"  {len(pendientes)} leads cambiaron de etapa despues de FECHA_HASTA; "
        "buscando esos movimientos...")
    _avisar(0.93, "Buscando cambios de etapa posteriores a FECHA_HASTA...")
    posteriores = descargar_eventos(hasta + 1, ahora, avance=(0.93, 0.95))
    nuevas = armar_filas([e for e in posteriores if e["entity_id"] in pendientes],
                         etapas, usuarios, embudos)
    filas = filas + nuevas
    return filas, {f["LEAD_ID"] for f in nuevas}, leads_que_no_cuadran(filas, tarjetas)


# ----------------------------------------------------------------------
# SELECCION POR ETIQUETA Y ANOMALIAS
# ----------------------------------------------------------------------
def etiquetas_con_texto(etiquetas, texto=None):
    """Etiquetas que contienen `texto` (sin importar mayúsculas ni acentos)."""
    buscado = normalizar(texto or ETIQUETA_LEADS)
    return [e for e in etiquetas if buscado in normalizar(e)]


def seleccionar_leads(tarjetas):
    """Leads con al menos una etiqueta que contiene ETIQUETA_LEADS."""
    return {lid for lid, t in tarjetas.items()
            if etiquetas_con_texto(t.get("lista_etiquetas") or [])}


# Anomalías que se reportan de los leads del historial (con etiqueta CIERRES)
ANOMALIAS = (
    ("varias_etiquetas", 'Más de una etiqueta con "{etiqueta}"'),
    ("sin_fecha_dictamen", "Sin FECHA DICTAMEN"),
    ("sin_estatus", "Sin ESTATUS DE NEGOCIO"),
    ("etapa_distinta", "Su etapa actual en Kommo no es la última del historial "
                       "(p. ej. se movió a otro embudo)"),
)
# Resultado de la última revisión: {clave: [LEAD_IDs]} (lo usa app.py)
ULTIMAS_ANOMALIAS = {}
# Archivo con las anomalías; app.py lo deja junto al historial descargado.
NOMBRE_ARCHIVO_ANOMALIAS = "leads_anomalos.txt"


def guardar_anomalias(carpeta, fecha_desde, fecha_hasta):
    """Escribe el texto de las anomalías en <carpeta>/leads_anomalos.txt.

    Si el archivo ya existe (misma carpeta de periodo), se reemplaza, igual
    que el historial. Devuelve la ruta, o None si no se pudo escribir.
    """
    def dia(fecha):
        return datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y") if fecha else "hoy"

    ruta = Path(carpeta) / NOMBRE_ARCHIVO_ANOMALIAS
    texto = (f"Descarga del {datetime.now():%d/%m/%Y %H:%M} | "
             f"Periodo: {dia(fecha_desde)} al {dia(fecha_hasta)}\n\n"
             f"{texto_anomalias(ULTIMAS_ANOMALIAS)}\n")
    try:
        # utf-8-sig: el Bloc de notas de Windows muestra bien los acentos
        ruta.write_text(texto, encoding="utf-8-sig")
    except OSError as e:
        print(f"Aviso: no se pudo guardar {ruta}: {e}")
        return None
    print(f"Anomalías guardadas en: {ruta}")
    return ruta


def revisar_anomalias(lead_ids, tarjetas, etapa_distinta=()):
    """{clave de ANOMALIAS: [LEAD_IDs ordenados]} de los leads indicados.

    etapa_distinta: leads cuyo embudo/etapa actual no es el del historial.
    """
    grupos = {clave: [] for clave, _ in ANOMALIAS}
    grupos["etapa_distinta"] = sorted(etapa_distinta)
    for lid in sorted(lead_ids):
        tarjeta = tarjetas.get(lid, {})
        campos = tarjeta.get("campos", {})
        if len(etiquetas_con_texto(tarjeta.get("lista_etiquetas") or [])) > 1:
            grupos["varias_etiquetas"].append(lid)
        if campos.get(normalizar("FECHA DICTAMEN")) in (None, ""):
            grupos["sin_fecha_dictamen"].append(lid)
        if campos.get(normalizar("ESTATUS DE NEGOCIO")) in (None, ""):
            grupos["sin_estatus"].append(lid)
    return grupos


# ----------------------------------------------------------------------
# LEADS CON VARIAS ETIQUETAS CIERRES: se asignan al asesor de la etiqueta
# puesta mas recientemente
# ----------------------------------------------------------------------
# {LEAD_ID: {"etiqueta": elegida o None, "fecha": datetime, "ignoradas": [...],
#            "etiquetas": [todas sus etiquetas CIERRES]}}  (texto_anomalias / app.py)
ULTIMAS_ASIGNACIONES = {}


def eventos_de_etiquetas(lead_ids):
    """Eventos 'etiqueta agregada' de esos leads, SIN limite de fechas (la
    etiqueta pudo ponerse antes de FECHA_DESDE). Se piden de 10 en 10."""
    eventos, ids = [], sorted(lead_ids)
    for i in range(0, len(ids), 10):
        params = {"filter[entity][]": "lead", "filter[type][0]": "entity_tag_added",
                  "filter[entity_id][]": ids[i:i + 10], "limit": 100, "page": 1}
        while True:
            data = get(f"{BASE}/events", params)
            if not data:
                break
            eventos.extend(data["_embedded"]["events"])
            if not data.get("_links", {}).get("next"):
                break
            params["page"] += 1
    return eventos


def asignar_etiqueta_mas_reciente(tarjetas, lead_ids):
    """Para cada lead con varias etiquetas CIERRES, elige la puesta más
    recientemente (entre las que tiene HOY). Si no se puede decidir (se
    pusieron en el mismo momento, o Kommo no tiene los eventos), no se elige.

    Una etiqueta sin evento se considera más antigua que las que sí lo tienen
    (se puso antes de que Kommo guardara el historial, o al crear el lead).
    """
    ultima = {}                        # (LEAD_ID, etiqueta normalizada) -> timestamp
    for ev in eventos_de_etiquetas(lead_ids):
        for item in ev.get("value_after") or []:
            nombre = ((item or {}).get("tag") or {}).get("name")
            if nombre:
                clave = (ev["entity_id"], normalizar(nombre))
                ultima[clave] = max(ultima.get(clave, 0), ev["created_at"])

    asignaciones = {}
    for lid in sorted(lead_ids):
        cierres = etiquetas_con_texto(tarjetas[lid].get("lista_etiquetas") or [])
        fechas = {t: ultima.get((lid, normalizar(t))) for t in cierres}
        conocidas = {t: f for t, f in fechas.items() if f is not None}
        mas_reciente = max(conocidas.values(), default=None)
        ganadoras = [t for t, f in conocidas.items() if f == mas_reciente]
        if len(ganadoras) == 1:
            elegida = ganadoras[0]
            asignaciones[lid] = {
                "etiqueta": elegida, "etiquetas": cierres,
                "fecha": datetime.fromtimestamp(mas_reciente, TZ).replace(tzinfo=None),
                "ignoradas": [t for t in cierres if t != elegida]}
        else:
            asignaciones[lid] = {"etiqueta": None, "etiquetas": cierres,
                                 "fecha": None, "ignoradas": []}
    return asignaciones


def aplicar_asignaciones(tarjetas, asignaciones):
    """Deja en el texto de ETIQUETAS solo la etiqueta CIERRES elegida (y las
    que no son CIERRES), para que el lead solo salga en el reporte de ese
    asesor. La lista original (lista_etiquetas) no se toca."""
    for lid, a in asignaciones.items():
        if a["etiqueta"]:
            quedan = [t for t in tarjetas[lid].get("lista_etiquetas") or []
                      if t not in a["ignoradas"]]
            tarjetas[lid]["etiquetas"] = ", ".join(quedan)


def _detalle_asignaciones(ids):
    """Renglones con a quién se asignó cada lead con varias etiquetas CIERRES."""
    lineas = []
    for lid in ids:
        a = ULTIMAS_ASIGNACIONES.get(lid)
        if not a:
            continue
        if a["etiqueta"]:
            lineas.append(f"    {lid} → {a['etiqueta']} (etiqueta puesta el "
                          f"{a['fecha']:%d/%m/%Y %H:%M}); se ignoró: {', '.join(a['ignoradas'])}")
        else:
            lineas.append(f"    {lid} → no se pudo determinar el asesor; cuenta para: "
                          f"{', '.join(a['etiquetas'])}")
    return lineas


def texto_anomalias(grupos):
    """Texto para la consola / ventana con los LEAD_IDs agrupados por anomalía."""
    lineas = []
    for clave, titulo in ANOMALIAS:
        ids = grupos.get(clave) or []
        if ids:
            titulo = titulo.format(etiqueta=ETIQUETA_LEADS)
            lineas.append(f"  {titulo} ({len(ids)}): {', '.join(map(str, ids))}")
            if clave == "varias_etiquetas":
                lineas.extend(_detalle_asignaciones(ids))
    if not lineas:
        return f'Leads con etiqueta "{ETIQUETA_LEADS}": sin anomalías.'
    return (f'Leads con etiqueta "{ETIQUETA_LEADS}" que requieren revisión:\n'
            + "\n".join(lineas))


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


def descargar_eventos(desde=None, hasta=None, avance=(0.12, 0.70)):
    """Parte el rango en HILOS tramos y los baja en paralelo.

    Sin argumentos usa FECHA_DESDE y FECHA_HASTA. desde / hasta: timestamps.
    avance: (inicio, fin) de la barra de progreso para este tramo.
    """
    if desde is None and hasta is None:
        desde = a_timestamp(FECHA_DESDE)
        hasta = a_timestamp(FECHA_HASTA, fin_de_dia=True)
    inicio_avance, fin_avance = avance

    if desde is None or HILOS <= 1:
        eventos = _descargar_rango(desde, hasta)
        _avisar(fin_avance, f"Eventos descargados: {len(eventos)}")
        return eventos

    fin = hasta or int(time.time())
    paso = max(1, (fin - desde) // HILOS)
    rangos, cursor = [], desde
    while cursor < fin:
        rangos.append((cursor, min(cursor + paso - 1, fin)))
        cursor += paso

    eventos = []
    with ThreadPoolExecutor(max_workers=HILOS) as ex:
        for hechos, lote in enumerate(ex.map(lambda r: _descargar_rango(*r), rangos), start=1):
            eventos.extend(lote)
            _avisar(inicio_avance + (fin_avance - inicio_avance) * hechos / len(rangos),
                    f"Descargando eventos... {len(eventos)} encontrados")

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
def armar_filas(eventos, etapas, usuarios, embudos):
    """Un renglón por cambio de etapa dentro de los embudos configurados
    (cualquier etapa). El filtro por etiqueta se aplica después, con las
    tarjetas de los leads."""
    filas = []
    for ev in eventos:
        sid_ant, pid_ant = extraer_status(ev.get("value_before"))
        sid_new, pid_new = extraer_status(ev.get("value_after"))
        pipeline_id = pid_new or pid_ant

        cfg = embudos.get(pipeline_id)
        if cfg is None:          # embudo que no nos interesa
            continue

        lead_id = ev["entity_id"]

        _, _, eta_new, orden = buscar_etapa(etapas, pid_new, sid_new)
        _, emb_ant, eta_ant, _ = buscar_etapa(etapas, pid_ant, sid_ant)

        fecha = datetime.fromtimestamp(ev["created_at"], TZ).replace(
            tzinfo=None)

        fila = {
            "LEAD_ID": lead_id,
            "EMBUDO": cfg["nombre_kommo"],
            "PIPELINE_ID": pipeline_id,   # auxiliar: no se exporta
            "STATUS_ID": sid_new,         # auxiliar: etapa nueva (id)
            "EMBUDO_ANTERIOR": emb_ant,   # auxiliar: embudo de la etapa anterior
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
    """Orden de las columnas de la hoja HISTORIAL (nombres internos).

    Los campos de la tarjeta (CAMPOS_TARJETA) van despues de la fecha de
    creacion, porque son datos del lead y no del movimiento. En el Excel se
    titulan segun ENCABEZADOS_HISTORIAL. MOVIDO_POR (INCLUIR_USUARIOS) va
    al final para no alterar el orden de las demas.
    """
    return (["LEAD_ID", "EMBUDO", COL_CREACION] + list(CAMPOS_TARJETA) +
            ([COL_ETIQUETAS] if INCLUIR_ETIQUETAS else []) +
            ["ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA",
             "DIAS_EN_ETAPA_ANTERIOR", "MOVIDO_POR"])


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
# EXCEL
# ----------------------------------------------------------------------
def escribir_excel(ruta, df, etiquetas=None):
    os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
    if INCLUIR_ETIQUETAS:
        df = marcar_etiquetas(df, etiquetas or {})
    historial = (df[[c for c in cols_historial() if c in df.columns]]
                 .rename(columns=ENCABEZADOS_HISTORIAL))
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

        # Montos con formato de dinero (las celdas vacias quedan vacias)
        ws = xl.sheets["HISTORIAL"]
        for idx, celda in enumerate(ws[1], start=1):
            if celda.value in CAMPOS_DINERO:
                for (c,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                    if isinstance(c.value, (int, float)):
                        c.number_format = FORMATO_DINERO
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
          f"Embudos: {', '.join(e['NOMBRE'] for e in EMBUDOS_CFG)} (todas las etapas) | "
          f"Etiqueta: {ETIQUETA_LEADS} | Archivos separados: {ARCHIVOS_SEPARADOS}")
    if not TOKEN:
        sys.exit("ERROR: falta el token de Kommo (variable de entorno "
                 "KOMMO_TOKEN o secret.json -> kommo -> TOKEN)")

    global ULTIMAS_ANOMALIAS, ULTIMAS_ASIGNACIONES
    ULTIMAS_ANOMALIAS = {}
    ULTIMAS_ASIGNACIONES = {}
    inicio = time.time()
    _prog["ev"] = 0
    _avisar(0.02, "Conectando con Kommo...")

    log("Cargando catalogo de embudos y etapas...")
    etapas, nombres_embudo = cargar_catalogo_etapas()
    embudos = resolver_embudos(etapas, nombres_embudo)
    usuarios = cargar_usuarios()
    _avisar(0.07, "Revisando las etapas de los embudos...")

    log("Descargando eventos...")
    _avisar(0.12, "Descargando eventos...")
    eventos = descargar_eventos()
    if not eventos:
        sys.exit("No se encontraron eventos en el rango solicitado.")

    filas = armar_filas(eventos, etapas, usuarios, embudos)
    if not filas:
        sys.exit("No hubo movimientos en los embudos configurados en el rango solicitado.")

    # Tarjetas de los leads con movimientos: fecha de creacion, campos
    # personalizados y etiquetas. Con las etiquetas se eligen los leads.
    log("Leyendo tarjetas de los leads...")
    _avisar(0.75, "Leyendo tarjetas de los leads...")
    tarjetas = cargar_tarjetas({f["LEAD_ID"] for f in filas})

    seleccionados = seleccionar_leads(tarjetas)
    filas = [f for f in filas if f["LEAD_ID"] in seleccionados]
    log(f"  {len(seleccionados)} leads con etiqueta que contiene '{ETIQUETA_LEADS}'")
    if not filas:
        sys.exit(f"Ningun lead con etiqueta que contenga '{ETIQUETA_LEADS}' tuvo "
                 "movimientos en el rango solicitado.")

    # Leads con varias etiquetas CIERRES: se asignan al asesor de la etiqueta
    # puesta más recientemente (en ETIQUETAS solo queda esa)
    varias = {lid for lid in {f["LEAD_ID"] for f in filas}
              if len(etiquetas_con_texto(tarjetas[lid].get("lista_etiquetas") or [])) > 1}
    if varias:
        _avisar(0.93, "Revisando leads con varias etiquetas de asesor...")
        ULTIMAS_ASIGNACIONES = asignar_etiqueta_mas_reciente(tarjetas, varias)
        aplicar_asignaciones(tarjetas, ULTIMAS_ASIGNACIONES)
    etiquetas = {lid: t["etiquetas"] for lid, t in tarjetas.items()}

    filas, descartados = aplicar_campos(filas, tarjetas, embudos)
    if descartados:
        log(f"  {descartados} leads descartados por campos vacios")
    if not filas:
        sys.exit("Ningun lead tiene llenos todos los CAMPOS_REQUERIDOS "
                 f"({', '.join(CAMPOS_REQUERIDOS) or 'sin campos'}). "
                 "Revisa que los nombres coincidan con los de Kommo.")

    # Que el historial de cada lead termine en su etapa actual en Kommo
    filas, completados, no_cuadran = completar_estado_actual(
        filas, tarjetas, etapas, usuarios, embudos)
    if completados:
        print(f"{len(completados)} leads cambiaron de etapa después de FECHA_HASTA "
              "(p. ej. de VENTAS a CIERRES): se agregaron esos movimientos para "
              "reflejar su etapa actual.")

    df = construir_df(filas)
    ULTIMAS_ANOMALIAS = revisar_anomalias(set(df["LEAD_ID"]), tarjetas, no_cuadran)
    _avisar(0.95, "Guardando el archivo...")

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
    print(texto_anomalias(ULTIMAS_ANOMALIAS))
    _avisar(1.0, "Descarga terminada")


def descargar_historial(fecha_desde, fecha_hasta, salida, al_avanzar=None):
    """Descarga el historial para app.py y lo guarda en `salida`.

    fecha_desde, fecha_hasta : textos 'AAAA-MM-DD' (el día final se incluye completo).
    salida                   : ruta del Excel. Siempre es un solo archivo, aunque
                               configuration.json tenga ARCHIVOS_SEPARADOS.
    al_avanzar               : función(fraccion, texto) para mostrar el avance.

    El archivo se escribe primero con otro nombre y solo reemplaza al anterior
    si la descarga termina bien: si algo falla, el historial anterior queda
    intacto. Si falla, lanza RuntimeError con el motivo.
    """
    global FECHA_DESDE, FECHA_HASTA, SALIDA, ARCHIVOS_SEPARADOS, AL_AVANZAR
    salida = Path(salida)
    temporal = salida.with_name(f"{salida.stem}.descargando{salida.suffix}")
    anteriores = (FECHA_DESDE, FECHA_HASTA, SALIDA, ARCHIVOS_SEPARADOS, AL_AVANZAR)
    FECHA_DESDE, FECHA_HASTA = fecha_desde, fecha_hasta
    SALIDA, ARCHIVOS_SEPARADOS, AL_AVANZAR = str(temporal), False, al_avanzar
    try:
        try:
            main()
        except SystemExit as e:        # main() termina con sys.exit("motivo")
            raise RuntimeError(str(e.code) if e.code else "La descarga se detuvo.") from None
        os.replace(temporal, salida)   # reemplaza por completo al anterior
        guardar_anomalias(salida.parent, fecha_desde, fecha_hasta)
        return salida
    finally:
        FECHA_DESDE, FECHA_HASTA, SALIDA, ARCHIVOS_SEPARADOS, AL_AVANZAR = anteriores
        if temporal.exists():
            try:
                temporal.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    main()