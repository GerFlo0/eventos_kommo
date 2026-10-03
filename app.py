"""
Programa para generar los acumulados, descargar el historial de etapas de
Kommo y generar los reportes individuales y el reporte general.

Flujo en la ventana:
  1. Elegir FECHA_DESDE y FECHA_HASTA (se recuerdan; la primera vez salen de
     configuration.json).
  2. "Generar acumulados": descarga la base de SharePoint
     (descargar_base_acumulados.py; enlaces en secret.json -> sharepoint) y
     genera ACUMULADOS con ese periodo (generar_acumulados.py). El periodo queda
     anotado junto a ACUMULADOS y es el que usan los reportes (carpeta, título y
     "Ventas del día"). Si algo falla, se detiene y lo explica.
  3. Elegir la carpeta de reportes. Cada paquete de reportes queda en
     <carpeta>/<FECHA_HASTA>/GENERALES/<prefijo>reporte_general_...xlsx y
     <carpeta>/<FECHA_HASTA>/INDIVIDUALES/<prefijo>/<prefijo>reporte_individual_...xlsx
  4. "Descargar historial": tarjetas, etiquetas y movimientos (en cualquier embudo,
     hasta el momento de la descarga) de
     todos y únicamente los leads de ACUMULADOS. Se guardan todos; los reportes
     solo usan los completos (con FECHA DICTAMEN, ESTATUS DE NEGOCIO y etiqueta
     de un asesor). Los incompletos se listan en leads_anomalos.txt.
     Selector de eventos (no se guarda; siempre arranca en el primero):
       - "Hasta el momento de ejecución": todos los eventos; el tiempo transcurrido
         y los días en la última etapa se miden hasta ahora.
       - "Solo hasta el corte": solo los eventos hasta FECHA_HASTA a las 23:59:59;
         el tiempo se mide hasta ese momento y las etiquetas se evalúan al corte
         (una etiqueta puesta después no cuenta; un lead sin asesor al corte no sale).
  5-6. (opcional) Estatus de negocio y descripción del reporte general.
  7. Prefijo del paquete (obligatorio): identifica el paquete (p. ej.
     "sin_restringido_" o "solo_restringido_") y nombra la subcarpeta de sus
     reportes individuales. Generar con un prefijo ya usado reemplaza ese paquete.
  8. "Generar reportes".

Rutas (configuration.json -> settings): CARPETA_BASE_ACUMULADOS,
RUTA_ACUMULADOS, CARPETA_HISTORIAL y CARPETA_ANOMALOS.
"""
import calendar
import contextlib
import glob
import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import duckdb as db
import pandas as pd
import requests

import acumulados as acum
import functions as fn
from general_report import NOMBRE_ARCHIVO_GENERAL, PREFIJO_INDIVIDUAL, generar_reporte_general
from individual_reports import (MESES, _nombre_archivo, aplicar_asignacion, cargar_asignaciones,
                                generar_reporte_estado_leads, normalizar,
                                normalizar_columnas_historial)

COL_ESTATUS = "ESTATUS DE NEGOCIO"
COL_DICTAMEN = "FECHA DICTAMEN"
# Ubicación de la lista de estatus dentro de secret.json. Ajusta las claves
# si la guardaste con otro nombre; p. ej. ["people", "estatus"] para
# secret.json -> people -> estatus.
CLAVES_LISTA_ESTATUS = ["estatus_negocio"]
# Así une el extractor los valores de un campo de opción múltiple.
SEPARADOR_MULTIPLE = " | "
# Caracteres que Windows no permite en nombres de archivo.
CARACTERES_PROHIBIDOS = set('<>:"/\\|?*')
FORMATO_FECHA_VISTA = "%d/%m/%Y"


# ---------------------------------------------------------------------------
# Lógica (sin interfaz)
# ---------------------------------------------------------------------------
def leer_lista_estatus(secret):
    """Lista de estatus de negocio guardada en secret.json."""
    nodo = secret
    for clave in CLAVES_LISTA_ESTATUS:
        if not isinstance(nodo, dict) or clave not in nodo:
            raise KeyError("No se encontró la lista de estatus en secret.json -> "
                           + " -> ".join(CLAVES_LISTA_ESTATUS)
                           + ". Ajusta CLAVES_LISTA_ESTATUS en app.py.")
        nodo = nodo[clave]
    if not isinstance(nodo, list) or not nodo:
        raise ValueError("La lista de estatus en secret.json está vacía o no es una lista.")
    return [str(e) for e in nodo]


def _estatus_de_celda(valor):
    """Estatus normalizados de una celda (puede traer varios: 'A | B')."""
    if valor is None or pd.isna(valor):
        return set()
    return {normalizar(p) for p in str(valor).split(SEPARADOR_MULTIPLE) if normalizar(p)}


def _estatus_por_lead(df):
    """{LEAD_ID: valor de ESTATUS DE NEGOCIO} (es el mismo en todas sus filas)."""
    return df.groupby("LEAD_ID")[COL_ESTATUS].first()


def estatus_no_listados(df, lista):
    """Valores de ESTATUS DE NEGOCIO del historial que no están en la lista."""
    conocidos = {normalizar(e) for e in lista}
    faltantes = set()
    for valor in _estatus_por_lead(df).dropna():
        for parte in str(valor).split(SEPARADOR_MULTIPLE):
            if normalizar(parte) and normalizar(parte) not in conocidos:
                faltantes.add(parte.strip())
    return sorted(faltantes)


def validar_historial(df):
    faltan = [c for c in ("LEAD_ID", COL_ESTATUS, COL_DICTAMEN) if c not in df.columns]
    if faltan:
        raise ValueError(f"Al historial le faltan las columnas {faltan}. "
                         "Descarga el historial de nuevo.")


def filtrar_historial(df, estatus, rango_dictamen=None):   # rango_dictamen: ya no se usa en la ventana
    """Historial completo de los leads que cumplen los filtros.

    estatus        : estatus de negocio elegidos (sin importar acentos ni
                     mayúsculas). Un lead entra si alguno de sus estatus
                     está en la lista.
    rango_dictamen : (desde, hasta) como date; si se indica, solo entran los
                     leads cuya FECHA DICTAMEN cae entre esas dos fechas,
                     incluidas. Los leads sin FECHA DICTAMEN quedan fuera.
    """
    elegidos = {normalizar(e) for e in estatus}
    por_lead = _estatus_por_lead(df)
    leads = set(por_lead[por_lead.map(lambda v: bool(_estatus_de_celda(v) & elegidos))].index)
    if rango_dictamen is not None:
        desde, hasta = (pd.Timestamp(f) for f in rango_dictamen)
        dictamen = pd.to_datetime(df.groupby("LEAD_ID")[COL_DICTAMEN].first(),
                                  errors="coerce").dt.normalize()   # solo el día
        leads &= set(dictamen[(dictamen >= desde) & (dictamen <= hasta)].index)
    return df[df["LEAD_ID"].isin(leads)].copy()


# Subcarpetas fijas dentro de <carpeta de reportes>/<fecha de corte>/
CARPETA_GENERALES = "GENERALES"
CARPETA_INDIVIDUALES_PAQUETE = "INDIVIDUALES"
# Nombres que Windows no permite como carpeta o archivo
NOMBRES_RESERVADOS = {"con", "prn", "aux", "nul",
                      *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}


def validar_prefijo(texto):
    """Prefijo de un paquete de reportes (obligatorio). Va al inicio del nombre
    de cada reporte y es el nombre de la subcarpeta de sus reportes individuales,
    así que debe ser un nombre de carpeta válido en Windows."""
    texto = (texto or "").strip()
    if not texto:
        raise ValueError("Escribe un prefijo para los reportes (p. ej. sin_restringido_).")
    prohibidos = sorted(set(texto) & CARACTERES_PROHIBIDOS)
    if prohibidos:
        raise ValueError(f"Caracteres no permitidos en nombres de archivo: {' '.join(prohibidos)}")
    if texto.endswith("."):
        raise ValueError("El prefijo no puede terminar en punto; usa _ como separador.")
    if texto.lower() in NOMBRES_RESERVADOS:
        raise ValueError(f'"{texto}" es un nombre reservado de Windows; usa otro prefijo.')
    return texto


def nombre_carpeta_reportes(fecha):
    """Subcarpeta de los reportes de un periodo: FECHA_HASTA como AAAA-MM-DD."""
    return f"{fecha:%Y-%m-%d}"


def ruta_historial():
    """Ubicación del historial: configuration.json -> CARPETA_HISTORIAL.
    Todo lo que lee, escribe o describe el historial usa esta función."""
    return fn.ruta_historial()


def carpetas_del_paquete(carpeta_reportes, fecha, prefijo):
    """(carpeta del día, carpeta de individuales, carpeta de generales) de un paquete:
    <carpeta>/<AAAA-MM-DD>/INDIVIDUALES/<prefijo>/ y <carpeta>/<AAAA-MM-DD>/GENERALES/."""
    dia = Path(carpeta_reportes) / nombre_carpeta_reportes(fecha)
    return dia, dia / CARPETA_INDIVIDUALES_PAQUETE / prefijo, dia / CARPETA_GENERALES


# Modos de eventos (selector del paso 4; no se guarda: siempre arranca en "ejecución")
MODO_EJECUCION = "ejecucion"   # todos los eventos hasta el momento de ejecución (por defecto)
MODO_CORTE = "corte"           # solo los eventos hasta FECHA_HASTA a las 23:59:59


def fin_del_dia(fecha):
    """FECHA_HASTA a las 23:59:59."""
    return datetime.combine(fecha, datetime.max.time()).replace(microsecond=0)


def cargar_etiquetas(carpeta):
    """{LEAD_ID: {"asesores": {asesor: fecha de la etiqueta o None}, "buzon": nombre}}
    de leads_anomalos.json ("etiquetas_asesores"). Sin datos: {}."""
    try:
        datos = json.loads((Path(carpeta) / "leads_anomalos.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    etiquetas = {}
    for d in datos.get("etiquetas_asesores") or []:
        etiquetas[int(d["lead"])] = {
            "buzon": d.get("buzon"),
            "asesores": {a["asesor"]: (datetime.fromisoformat(a["fecha_etiqueta"])
                                       if a.get("fecha_etiqueta") else None)
                         for a in d.get("asesores") or []}}
    return etiquetas


def permitidos_al_corte(etiquetas, corte):
    """{LEAD_ID: asesores en cuyos reportes puede salir el lead}, evaluando las
    etiquetas al corte: una etiqueta puesta después del corte no cuenta (una sin
    fecha se considera anterior). Con un solo asesor al corte, sale solo con él;
    con varios, se asigna como siempre (buzón, luego la etiqueta más reciente
    al corte; en empate, con todos los de ese momento); sin ninguno, el lead no
    tenía asesor al corte y no sale en ningún reporte."""
    import extract_data_from_kommo as extractor     # misma regla del buzón que la descarga
    permitidos = {}
    for lid, d in etiquetas.items():
        presentes = {a: f for a, f in d["asesores"].items() if f is None or f <= corte}
        if len(presentes) <= 1:
            permitidos[lid] = set(presentes)
            continue
        por_buzon = extractor.asesor_del_buzon(d.get("buzon"), list(presentes))
        conocidas = {a: f for a, f in presentes.items() if f}
        mas_reciente = max(conocidas.values(), default=None)
        ganadores = [a for a, f in conocidas.items() if f == mas_reciente]
        if por_buzon:
            permitidos[lid] = {por_buzon}
        elif len(ganadores) == 1:
            permitidos[lid] = {ganadores[0]}
        else:
            permitidos[lid] = set(presentes)
    return permitidos


def aplicar_permitidos(historial, asesor, permitidos):
    """Quita del historial de un asesor los leads que al corte no eran suyos."""
    if not permitidos or historial is None or historial.empty:
        return historial
    clave = normalizar(asesor)
    ajenos = {lid for lid, asesores in permitidos.items()
              if clave not in {normalizar(a) for a in asesores}}
    return historial[~historial["LEAD_ID"].isin(ajenos)]


def aplicar_corte(df, corte):
    """Solo los movimientos hasta el corte (incluido).

    En el historial, ETIQUETAS viene solo en el último movimiento de cada lead;
    si ese movimiento queda después del corte, las etiquetas pasan al último
    movimiento que sí queda, para que la consulta de cada asesor siga
    encontrando al lead."""
    if df is None or corte is None:
        return df
    fechas = pd.to_datetime(df["FECHA"])
    cortado = df[fechas <= corte].copy()
    if "ETIQUETAS" in df.columns and not cortado.empty:
        por_lead = df.dropna(subset=["ETIQUETAS"]).groupby("LEAD_ID")["ETIQUETAS"].last()
        cortado["ETIQUETAS"] = None
        orden = cortado.assign(_FECHA=pd.to_datetime(cortado["FECHA"])).sort_values(["LEAD_ID", "_FECHA"])
        ultimos = orden.groupby("LEAD_ID").tail(1).index
        cortado.loc[ultimos, "ETIQUETAS"] = cortado.loc[ultimos, "LEAD_ID"].map(por_lead).values
    return cortado


def generar_reportes(df, secret, fecha, descripcion="", prefijo="", carpeta_reportes=None,
                     asignaciones=None, corte=None, etiquetas=None):
    """Reportes individuales de cada asesor + reporte general.

    fecha            : FECHA_HASTA; nombra la carpeta y los archivos.
    prefijo          : obligatorio. Identifica el paquete de reportes: va al inicio
                       del nombre de cada reporte y nombra la subcarpeta de sus
                       individuales. Estructura:
                         <carpeta_reportes>/<AAAA-MM-DD>/GENERALES/<prefijo>reporte_general_...
                         <carpeta_reportes>/<AAAA-MM-DD>/INDIVIDUALES/<prefijo>/<prefijo>reporte_individual_...
    carpeta_reportes : carpeta elegida en la ventana (si no se indica,
                       tablas/reportes dentro de la carpeta de datos).
    corte            : None (modo por defecto) = todos los eventos hasta hoy, y el
                       tiempo transcurrido se mide hasta el momento de ejecución.
                       Fecha y hora (FECHA_HASTA 23:59:59) = solo los eventos hasta
                       ese momento, el tiempo se mide hasta ahí y las etiquetas se
                       evalúan al corte (etiquetas: de cargar_etiquetas).
    asignaciones     : {LEAD_ID: asesor} de los leads con varios asesores (de
                       leads_anomalos.json); cada uno sale solo en el reporte
                       de su asesor. Si no se indica, se lee de CARPETA_ANOMALOS.
    Devuelve la ruta del reporte general.
    """
    prefijo = validar_prefijo(prefijo)
    _, carpeta_individuales, carpeta_general = carpetas_del_paquete(
        carpeta_reportes or fn.ruta_configurada("CARPETA_REPORTES", "tablas/reportes"), fecha, prefijo)

    # El reporte general junta los reportes individuales de la subcarpeta del
    # paquete: se borran los de una corrida anterior del mismo paquete para no
    # mezclar filtros (los de otros prefijos están en otras subcarpetas).
    if carpeta_individuales.is_dir():
        for viejo in carpeta_individuales.glob(f"{glob.escape(prefijo)}{PREFIJO_INDIVIDUAL}*.xlsx"):
            viejo.unlink()   # PermissionError si está abierto en Excel

    if asignaciones is None:
        asignaciones = cargar_asignaciones(fn.carpeta_anomalos())
    permitidos = None
    if corte is not None:
        df = aplicar_corte(df, corte)
        if etiquetas is None:
            etiquetas = cargar_etiquetas(fn.carpeta_anomalos())
        if etiquetas:
            permitidos = permitidos_al_corte(etiquetas, corte)
        else:
            print("Aviso: el historial no tiene las fechas de las etiquetas; los leads con "
                  "varios asesores se asignan como en el modo por defecto. Vuelve a descargarlo.")
        fecha_corte = corte                      # referencia del tiempo transcurrido
    else:
        fecha_corte = datetime.now().replace(microsecond=0)  # misma para todos los reportes
    generados = []
    con = db.connect()
    con.register("df", df)   # la consulta de secret.json lee la tabla "df"
    try:
        for asesor in secret["asesores"]:
            resultado = con.execute(secret["query"], [f"%{asesor}%"]).df()
            if permitidos is not None:
                resultado = aplicar_permitidos(resultado, asesor, permitidos)
            else:
                resultado = aplicar_asignacion(resultado, asesor, asignaciones)
            ruta = generar_reporte_estado_leads(
                resultado, asesor=asesor, fecha_corte=fecha_corte,
                ruta_salida=carpeta_individuales / _nombre_archivo(asesor, prefijo))
            if ruta:
                generados.append(ruta)
    finally:
        con.close()

    if not generados:
        raise ValueError("Ningún asesor tiene leads con los filtros elegidos; "
                         "no se generó el reporte general.")
    return generar_reporte_general(fecha=fecha, carpeta_individuales=carpeta_individuales,
                                   carpeta_salida=carpeta_general,
                                   descripcion=descripcion, prefijo=prefijo,
                                   prefijo_individuales=prefijo)


# ---- Historial de etapas (en la carpeta del periodo) ------------------------
def ruta_info_historial(ruta_historial):
    """Datos de la última descarga, junto al historial (…historial_etapas_kommo.info.json)."""
    ruta = Path(ruta_historial)
    return ruta.with_name(f"{ruta.stem}.info.json")


def guardar_info_historial(ruta_historial, historial):
    """Anota cuándo se descargó (creó) el historial."""
    info = {
        "descargado": datetime.now().isoformat(timespec="seconds"),
        "leads": int(historial["LEAD_ID"].nunique()),
        "movimientos": int(len(historial)),
        # Si el historial se reemplaza por otro medio (p. ej. corriendo el
        # extractor por consola), esta marca ya no coincide y la info se ignora.
        "marca_archivo": os.path.getmtime(ruta_historial),
    }
    fn._escribir_json_seguro(ruta_info_historial(ruta_historial), info)
    return info


def leer_info_historial(ruta_historial):
    """Info de la última descarga desde la ventana, o None si no hay o ya no corresponde."""
    ruta = Path(ruta_historial)
    try:
        info = json.loads(ruta_info_historial(ruta).read_text(encoding="utf-8"))
        if abs(float(info["marca_archivo"]) - os.path.getmtime(ruta)) > 1:
            return None
        info["descargado"] = datetime.fromisoformat(info["descargado"])
        return info
    except (OSError, ValueError, KeyError, TypeError):
        return None


def cargar_historial(ruta_historial):
    """(DataFrame, aviso). Si no hay historial o no sirve, DataFrame es None."""
    ruta = Path(ruta_historial)
    if not ruta.exists():
        return None, None
    try:
        # Los títulos del HISTORIAL ("LEAD ID", "FECHA EVENTO"...) se traducen a
        # los nombres que usan los reportes y la consulta de secret.json.
        historial = normalizar_columnas_historial(pd.read_excel(ruta))
        validar_historial(historial)
        return historial, None
    except Exception as e:  # noqa: BLE001 - archivo dañado o de una versión anterior
        return None, f"El historial guardado no se pudo usar ({e}). Descárgalo de nuevo."


def universo(historial, acumulados, asesores=None):
    """Historial de los leads de ACUMULADOS que están completos, con su ORIGEN
    (CONTACTACION / IA). Los demás leads del historial no entran en los reportes.

    asesores: si se indica, solo quedan los leads con la etiqueta de alguno
    (mismo criterio que la consulta LIKE de secret.json: el texto de ETIQUETAS
    contiene el nombre, distinguiendo mayúsculas)."""
    if historial is None or acumulados is None:
        return None
    origen = acumulados.set_index("LEAD_ID")["ORIGEN"]
    df = historial[historial["LEAD_ID"].isin(origen.index)].copy()
    # Solo los completos: con FECHA DICTAMEN (el ESTATUS lo exige el filtro de
    # estatus y la etiqueta de asesor, la consulta de secret.json).
    if COL_DICTAMEN in df.columns:
        con_dictamen = df.groupby("LEAD_ID")[COL_DICTAMEN].transform(lambda s: s.notna().any())
        df = df[con_dictamen]
    if asesores and "ETIQUETAS" in df.columns:
        etiquetas = df.groupby("LEAD_ID")["ETIQUETAS"].transform(
            lambda s: ", ".join(str(v) for v in s.dropna()))
        df = df[etiquetas.map(lambda texto: any(a in texto for a in asesores))]
    df["ORIGEN"] = df["LEAD_ID"].map(origen)
    return df


def acumulados_sin_historial(historial, acumulados, descartados=()):
    """LEAD_IDs de ACUMULADOS que no están en el historial (sin contar los que la
    descarga descartó a propósito por FECHA DICTAMEN posterior a FECHA_HASTA)."""
    if acumulados is None:
        return []
    en_historial = set(historial["LEAD_ID"]) if historial is not None else set()
    return sorted(set(acumulados["LEAD_ID"]) - en_historial - set(descartados))


def cargar_descartados(carpeta):
    """LEAD_IDs que la última descarga descartó por FECHA DICTAMEN posterior a
    FECHA_HASTA (leads_anomalos.json)."""
    try:
        datos = json.loads((Path(carpeta) / "leads_anomalos.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return set()
    bloque = datos.get("dictamen_posterior_al_corte") or {}
    return {int(d["lead"]) for d in bloque.get("leads") or []}


def cargar_acumulados():
    """(DataFrame, info, aviso). info = {"generado", "fecha_desde", "fecha_hasta"}:
    el periodo con que se generó (lo usan los reportes). Si no hay info (p. ej.
    se generó por fuera), el periodo sale de sus fechas. Si no se puede leer:
    (None, None, motivo)."""
    try:
        ruta = acum.ruta_acumulados()
        datos = acum.leer_acumulados(ruta)
        info = acum.leer_info(ruta)
        if info is None:
            desde, hasta = acum.periodo(datos)
            info = {"generado": None, "fecha_desde": desde, "fecha_hasta": hasta}
        return datos, info, None
    except Exception as e:  # noqa: BLE001 - se muestra en la ventana
        return None, None, str(e)


def cargar_datos():
    """Todo lo que la ventana necesita. Lanza una excepción solo si falta
    algo sin lo que no puede funcionar (secret.json); el historial es opcional
    porque se puede descargar desde la ventana."""
    secret = fn.import_json("json/secret.json")
    estatus = leer_lista_estatus(secret)
    prefs = fn.leer_preferencias()
    hoy = date.today()
    # Selector: las últimas fechas usadas; la primera vez, las de configuration.json
    fecha_desde = (_fecha_de_texto(prefs.get("fecha_desde")) or fn.fecha_desde()
                   or hoy.replace(day=1))
    fecha_hasta = _fecha_de_texto(prefs.get("fecha_hasta")) or fn.fecha_reportes()
    acumulados, info_acum, aviso_acum = cargar_acumulados()
    carpeta = prefs.get("carpeta_reportes") or None
    historial, aviso = cargar_historial(ruta_historial())   # el de CARPETA_HISTORIAL
    return {
        "secret": secret,
        "estatus": estatus,
        "historial": historial,
        "fecha": fecha_hasta,
        "fecha_desde": fecha_desde,
        "carpeta_reportes": carpeta,
        "aviso_inicial": aviso,
        "acumulados": acumulados,
        "info_acumulados": info_acum,
        "aviso_acumulados": aviso_acum,
    }


def _fecha_de_texto(valor):
    try:
        return date.fromisoformat(valor) if valor else None
    except (TypeError, ValueError):
        return None


def explicar_error_descarga(error):
    """Mensaje entendible para quien usa el programa, con el detalle técnico debajo."""
    codigo = None
    if isinstance(error, requests.exceptions.HTTPError) and error.response is not None:
        codigo = error.response.status_code
    if codigo == 401:
        motivo = ("Kommo rechazó el token (error 401): puede haber expirado o ser incorrecto. "
                  "Pide a quien administra el programa que lo actualice.")
    elif codigo == 403:
        motivo = ("Kommo negó el acceso (error 403): el token no tiene permiso o el "
                  "subdominio no es correcto.")
    elif codigo == 404:
        motivo = "No se encontró la cuenta de Kommo (error 404): revisa el subdominio."
    elif isinstance(error, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)) or \
            (isinstance(error, RuntimeError) and str(error).startswith("Fallaron")):
        motivo = ("No se pudo conectar con Kommo después de varios intentos. Revisa tu "
                  "conexión a internet e intenta de nuevo en unos minutos.")
    elif isinstance(error, PermissionError):
        motivo = ("No se pudo reemplazar el historial anterior; si está abierto en Excel, "
                  "ciérralo e intenta de nuevo.")
    else:
        return str(error) or type(error).__name__
    return f"{motivo}\n\nDetalle: {error}"


def abrir_en_explorador(ruta):
    """Abre una carpeta con el explorador de archivos del sistema."""
    if sys.platform.startswith("win"):
        os.startfile(ruta)  # noqa: S606 - solo en Windows
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


def carpeta_escribible(ruta):
    """True si se pueden crear archivos en la carpeta."""
    prueba = Path(ruta) / ".prueba_escritura_reportes"
    try:
        prueba.write_text("ok", encoding="utf-8")
        prueba.unlink()
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Calendario (tkinter puro: no necesita librerías extra, así compila sin
# problemas con PyInstaller)
# ---------------------------------------------------------------------------
MESES_TITULO = [m.capitalize() for m in MESES]
DIAS_SEMANA = ["Lu", "Ma", "Mi", "Ju", "Vi", "Sá", "Do"]


class Calendario(tk.Toplevel):
    """Ventanita con un mes para elegir una fecha."""

    def __init__(self, ancla, fecha, al_elegir):
        super().__init__(ancla)
        self.withdraw()
        self.title("Elegir fecha")
        self.resizable(False, False)
        self.transient(ancla.winfo_toplevel())
        self.al_elegir = al_elegir
        self.elegida = fecha
        self.mes = fecha.replace(day=1)

        cabecera = ttk.Frame(self, padding=(6, 6, 6, 0))
        cabecera.pack(fill="x")
        for texto, meses in (("«", -12), ("‹", -1)):
            ttk.Button(cabecera, text=texto, width=3,
                       command=lambda m=meses: self._mover(m)).pack(side="left")
        self.titulo = ttk.Label(cabecera, anchor="center", font=("TkDefaultFont", 10, "bold"))
        self.titulo.pack(side="left", fill="x", expand=True)
        for texto, meses in (("»", 12), ("›", 1)):
            ttk.Button(cabecera, text=texto, width=3,
                       command=lambda m=meses: self._mover(m)).pack(side="right")

        self.cuerpo = ttk.Frame(self, padding=6)
        self.cuerpo.pack()
        pie = ttk.Frame(self, padding=(6, 0, 6, 6))
        pie.pack(fill="x")
        ttk.Button(pie, text="Hoy", command=lambda: self._elegir(date.today())).pack(side="left")
        ttk.Button(pie, text="Cancelar", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda _e: self.destroy())
        self._dibujar()

        self.update_idletasks()
        x = ancla.winfo_rootx()
        y = ancla.winfo_rooty() + ancla.winfo_height() + 2
        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.grab_set()
        self.focus_set()

    def _mover(self, meses):
        total = self.mes.year * 12 + self.mes.month - 1 + meses
        self.mes = date(total // 12, total % 12 + 1, 1)
        self._dibujar()

    def _dibujar(self):
        for hijo in self.cuerpo.winfo_children():
            hijo.destroy()
        self.titulo.config(text=f"{MESES_TITULO[self.mes.month - 1]} {self.mes.year}")
        for col, dia in enumerate(DIAS_SEMANA):
            ttk.Label(self.cuerpo, text=dia, width=4, anchor="center").grid(row=0, column=col)
        hoy = date.today()
        semanas = calendar.Calendar(firstweekday=0).monthdatescalendar(self.mes.year, self.mes.month)
        for fila, semana in enumerate(semanas, start=1):
            for col, dia in enumerate(semana):
                boton = tk.Button(self.cuerpo, text=str(dia.day), width=3, relief="flat",
                                  command=lambda d=dia: self._elegir(d))
                if dia == self.elegida:
                    boton.config(bg="#1F3864", fg="white", activebackground="#305496",
                                 activeforeground="white")
                elif dia.month != self.mes.month:
                    boton.config(fg="gray60")
                if dia == hoy:
                    boton.config(font=("TkDefaultFont", 9, "bold underline"))
                boton.grid(row=fila, column=col, padx=1, pady=1)

    def _elegir(self, dia):
        self.al_elegir(dia)
        self.destroy()


class SelectorFecha(ttk.Frame):
    """Campo de fecha (dd/mm/aaaa) que se elige con un calendario."""

    def __init__(self, master, fecha, al_cambiar=None):
        super().__init__(master)
        self._fecha = fecha
        self.al_cambiar = al_cambiar
        self._habilitado = True
        self.texto = tk.StringVar(value=fecha.strftime(FORMATO_FECHA_VISTA))
        self.entrada = ttk.Entry(self, textvariable=self.texto, width=11, state="readonly")
        self.entrada.pack(side="left")
        self.entrada.bind("<Button-1>", lambda _e: self.abrir())
        self.boton = ttk.Button(self, text="▾", width=3, command=self.abrir)
        self.boton.pack(side="left", padx=(2, 0))

    def get(self):
        return self._fecha

    def set(self, fecha):
        cambio = fecha != self._fecha
        self._fecha = fecha
        self.texto.set(fecha.strftime(FORMATO_FECHA_VISTA))
        if cambio and self.al_cambiar:
            self.al_cambiar()

    def abrir(self):
        if self._habilitado:
            return Calendario(self, self._fecha, self.set)
        return None

    def habilitar(self, si):
        self._habilitado = si
        self.entrada.state(["!disabled", "readonly"] if si else ["disabled"])
        self.boton.state(["!disabled"] if si else ["disabled"])


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
class _Consola:
    """Manda lo que se imprime (print) al cuadro de mensajes de la ventana."""

    def __init__(self, ventana):
        self.ventana = ventana

    def write(self, texto):
        self.ventana.mensajes.insert("end", texto)
        self.ventana.mensajes.see("end")
        self.ventana.update_idletasks()

    def flush(self):
        pass


class _EscritorCola:
    """Lo que se imprime durante la descarga (en otro hilo) va a una cola que
    la ventana lee; tkinter no se puede tocar desde otro hilo."""

    def __init__(self, cola, ruta_final=None):
        self.cola = cola
        self.ruta_final = ruta_final

    def write(self, texto):
        if not texto or texto.startswith("\r"):      # "\r" = contador de avance de consola
            return len(texto or "")
        if self.ruta_final and ".descargando." in texto:
            # El extractor escribe primero un archivo temporal; se muestra la ruta final.
            _, _, resto = texto.partition(" | ")
            texto = f"Historial guardado en: {self.ruta_final} | {resto}"
        self.cola.put(("texto", texto))
        return len(texto)

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self, secret, estatus, historial, fecha, fecha_desde, carpeta_reportes=None,
                 aviso_inicial=None, acumulados=None, aviso_acumulados=None, info_acumulados=None):
        """fecha / fecha_desde: valores iniciales del selector. El periodo de los
        reportes (self.fecha_desde, self.fecha) es el del ACUMULADOS generado."""
        super().__init__()
        self.secret, self.estatus, self.historial = secret, estatus, historial
        self.sel_desde, self.sel_hasta = fecha_desde, fecha
        self.carpeta_reportes = carpeta_reportes
        self._fijar_acumulados(acumulados, info_acumulados, aviso_acumulados)
        self.etiquetas = cargar_etiquetas(fn.carpeta_anomalos())   # fechas de etiqueta (modo corte)
        self.descartados = cargar_descartados(fn.carpeta_anomalos())
        self._descargando = False
        self._generando_acumulados = False
        self._cola = None
        self.title("Reportes de leads - Kommo")
        self.minsize(980, 660)
        self._armar()
        self._actualizar_info_historial()
        self._actualizar_conteo()
        self._actualizar_estado()
        if aviso_inicial:
            self._escribir(f"Aviso: {aviso_inicial}\n")
        self._avisar_estatus_no_listados()
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

    # --- construcción ---------------------------------------------------------
    def _armar(self):
        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)
        marco.columnconfigure(0, weight=1, uniform="col")
        marco.columnconfigure(1, weight=1, uniform="col")
        marco.rowconfigure(1, weight=1)
        izquierda = ttk.Frame(marco)
        izquierda.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        derecha = ttk.Frame(marco)
        derecha.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # 1-2. Periodo y acumulados
        caja = ttk.LabelFrame(izquierda, text="1-2. Periodo y acumulados", padding=8)
        caja.pack(fill="x")
        fechas = ttk.Frame(caja)
        fechas.pack(fill="x")
        ttk.Label(fechas, text="Desde (FECHA_DESDE):").grid(row=0, column=0, sticky="w")
        self.selector_desde = SelectorFecha(fechas, self.sel_desde, self._al_cambiar_fechas)
        self.selector_desde.grid(row=0, column=1, sticky="w", padx=(6, 0), pady=(0, 3))
        ttk.Label(fechas, text="Hasta (FECHA_HASTA):").grid(row=1, column=0, sticky="w")
        self.selector_hasta = SelectorFecha(fechas, self.sel_hasta, self._al_cambiar_fechas)
        self.selector_hasta.grid(row=1, column=1, sticky="w", padx=(6, 0))
        self.boton_acumulados = ttk.Button(fechas, text="Generar acumulados",
                                           command=self._generar_acumulados)
        self.boton_acumulados.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.etapa_acumulados = ttk.Label(caja, foreground="gray40")
        self.etapa_acumulados.pack(anchor="w")
        self.info_acumulados = ttk.Label(caja, wraplength=440, justify="left")
        self.info_acumulados.pack(anchor="w")
        self.aviso_fechas = ttk.Label(caja, foreground="#B35C00", wraplength=440, justify="left")
        self.aviso_fechas.pack(anchor="w")

        # 3. Carpeta de reportes
        caja = ttk.LabelFrame(izquierda, text="3. Carpeta de reportes", padding=8)
        caja.pack(fill="x", pady=(10, 0))
        fila = ttk.Frame(caja)
        fila.pack(fill="x")
        self.texto_carpeta = tk.StringVar(value=self.carpeta_reportes or "")
        ttk.Entry(fila, textvariable=self.texto_carpeta, state="readonly", width=10).pack(
            side="left", fill="x", expand=True, padx=(0, 4))
        self.boton_carpeta = ttk.Button(fila, text="Elegir...", command=self._elegir_carpeta)
        self.boton_carpeta.pack(side="left")
        self.boton_abrir = ttk.Button(fila, text="Abrir", command=self._abrir_carpeta)
        self.boton_abrir.pack(side="left", padx=(4, 0))
        self.destino = ttk.Label(caja, foreground="gray40", wraplength=440)
        self.destino.pack(anchor="w", pady=(2, 0))

        # Historial de etapas
        caja = ttk.LabelFrame(izquierda, text="4. Historial de etapas (Kommo)", padding=8)
        caja.pack(fill="x", pady=(10, 0))
        self.info_historial = ttk.Label(caja, wraplength=440, justify="left")
        self.info_historial.pack(anchor="w")
        self.aviso_historial = ttk.Label(caja, foreground="#B35C00", wraplength=440, justify="left")
        self.aviso_historial.pack(anchor="w")
        fila = ttk.Frame(caja)
        fila.pack(fill="x", pady=(6, 0))
        self.barra = ttk.Progressbar(fila, mode="determinate", maximum=100)
        self.barra.pack(side="left", fill="x", expand=True)
        self.porcentaje = ttk.Label(fila, text="", width=5, anchor="e")
        self.porcentaje.pack(side="left", padx=(4, 8))
        self.boton_descargar = ttk.Button(fila, text="Descargar historial", command=self._descargar)
        self.boton_descargar.pack(side="left")
        self.etapa_descarga = ttk.Label(caja, foreground="gray40")
        self.etapa_descarga.pack(anchor="w")
        # Qué eventos usan los reportes (no se guarda: siempre arranca en "ejecución")
        ttk.Label(caja, text="Eventos a usar en los reportes (etapas, embudos y etiquetas):").pack(
            anchor="w", pady=(6, 0))
        self.modo_eventos = tk.StringVar(value=MODO_EJECUCION)
        ttk.Radiobutton(caja, text="Hasta el momento de ejecución (por defecto)",
                        variable=self.modo_eventos, value=MODO_EJECUCION,
                        command=self._al_cambiar_modo).pack(anchor="w")
        self.opcion_corte = ttk.Radiobutton(caja, variable=self.modo_eventos, value=MODO_CORTE,
                                            command=self._al_cambiar_modo)
        self.opcion_corte.pack(anchor="w")

        # Estatus de negocio (selección múltiple)
        caja = ttk.LabelFrame(derecha, text="5. Estatus de negocio a incluir (opcional)", padding=8)
        caja.pack(fill="x")
        lista_marco = ttk.Frame(caja)
        lista_marco.pack(fill="both", expand=True)
        self.lista = tk.Listbox(lista_marco, selectmode="multiple", exportselection=False,
                                height=min(max(len(self.estatus), 4), 6))
        barra = ttk.Scrollbar(lista_marco, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        for e in self.estatus:
            self.lista.insert("end", e)
        self.lista.select_set(0, "end")
        self.lista.bind("<<ListboxSelect>>", lambda _e: self._actualizar_conteo())
        botones = ttk.Frame(caja)
        botones.pack(anchor="w", pady=(6, 0))
        ttk.Button(botones, text="Todos", command=lambda: self._marcar(True)).pack(side="left")
        ttk.Button(botones, text="Ninguno", command=lambda: self._marcar(False)).pack(side="left", padx=6)

        # Descripción del reporte general
        caja = ttk.Frame(derecha)
        caja.pack(fill="x", pady=(10, 0))
        ttk.Label(caja, text="6. Descripción para el título del reporte general (opcional):").pack(anchor="w")
        self.descripcion = tk.StringVar()
        ttk.Entry(caja, textvariable=self.descripcion).pack(fill="x")

        # Texto al inicio de los nombres de archivo (vacío = nombre normal)
        # Prefijo del paquete de reportes (obligatorio): va al inicio de cada
        # nombre de archivo y nombra la subcarpeta de los individuales.
        caja = ttk.LabelFrame(derecha, text="7. Prefijo de los reportes (obligatorio)", padding=8)
        caja.pack(fill="x", pady=(10, 0))
        caja.columnconfigure(1, weight=1)
        self.prefijo = tk.StringVar()
        ttk.Label(caja, text="Prefijo:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(caja, textvariable=self.prefijo, width=10).grid(row=0, column=1, sticky="ew")
        self.vista_prefijo = ttk.Label(caja, foreground="gray40", wraplength=440, justify="left")
        self.vista_prefijo.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.prefijo.trace_add("write", lambda *_a: self._al_cambiar_prefijo())
        self._actualizar_nombres()

        # Conteo y botón
        fila = ttk.Frame(derecha)
        fila.pack(fill="x", pady=(10, 0))
        self.conteo = ttk.Label(fila, text="")
        self.conteo.pack(side="left")
        self.boton = ttk.Button(fila, text="8. Generar reportes", command=self._generar)
        self.boton.pack(side="right")
        self.requisito = ttk.Label(derecha, foreground="red", wraplength=440, justify="left")
        self.requisito.pack(anchor="w", pady=(4, 0))

        self.mensajes = ScrolledText(marco, height=8, wrap="word")
        self.mensajes.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

    # --- estado general -------------------------------------------------------
    def _ocupado(self):
        return self._descargando or self._generando_acumulados

    def _actualizar_estado(self):
        """Habilita o deshabilita los controles según lo que falte."""
        ocupado = self._ocupado()
        sin_carpeta = not self.carpeta_reportes
        sin_acumulados = self.acumulados is None
        fechas_mal = self.selector_desde.get() > self.selector_hasta.get()
        # El historial se descarga para los leads de ACUMULADOS: hace falta generarlos antes
        puede_descargar = not ocupado and not sin_acumulados
        sin_prefijo = self._prefijo_valido() is None
        puede_generar = (not ocupado and not sin_carpeta and not sin_acumulados
                         and self.historial is not None and not sin_prefijo)
        self.boton_acumulados.state(["disabled"] if (ocupado or fechas_mal) else ["!disabled"])
        self.boton_descargar.state(["!disabled"] if puede_descargar else ["disabled"])
        self.boton.state(["!disabled"] if puede_generar else ["disabled"])
        self.boton_carpeta.state(["disabled"] if ocupado else ["!disabled"])
        self.boton_abrir.state(["disabled"] if (sin_carpeta or ocupado) else ["!disabled"])
        for selector in (self.selector_desde, self.selector_hasta):
            selector.habilitar(not ocupado)
        self._actualizar_info_acumulados()
        if fechas_mal:
            requisito = "FECHA_DESDE no puede ser posterior a FECHA_HASTA."
        elif sin_acumulados and not ocupado:
            requisito = "Genera los acumulados (pasos 1-2) para continuar."
        elif self.historial is None and not ocupado:
            requisito = "Descarga el historial (paso 4) para poder generar reportes."
        elif sin_carpeta:
            requisito = "Elige la carpeta de reportes (paso 3) para poder generar reportes."
        elif sin_prefijo:
            requisito = "Escribe el prefijo de los reportes (paso 7) para poder generarlos."
        else:
            requisito = ""
        self.requisito.config(text=requisito)
        if self.carpeta_reportes:
            destino = Path(self.carpeta_reportes) / nombre_carpeta_reportes(self.fecha)
            self.destino.config(text=f"Los reportes se guardarán en: {destino}")
        else:
            self.destino.config(text="")

    @property
    def ruta_historial(self):
        """Ubicación del historial (configuration.json -> CARPETA_HISTORIAL)."""
        return ruta_historial()

    def _fijar_acumulados(self, acumulados, info, aviso):
        """Guarda el ACUMULADOS cargado. Su periodo es el de los reportes."""
        self.acumulados, self.info_acumulados_actual, self.aviso_acumulados = acumulados, info, aviso
        if info:
            self.fecha_desde, self.fecha = info["fecha_desde"], info["fecha_hasta"]
        else:                                   # sin ACUMULADOS: se muestra el selector
            self.fecha_desde, self.fecha = self.sel_desde, self.sel_hasta

    def _actualizar_info_acumulados(self):
        avisos = []
        if self.acumulados is None:
            self.info_acumulados.config(text=f"Aún no hay acumulados. Se guardarán en:\n"
                                             f"{acum.ruta_acumulados() or '(sin configurar)'}")
            if self.aviso_acumulados and "No se encontró" not in self.aviso_acumulados:
                avisos.append(f"No se pudo leer ACUMULADOS: {self.aviso_acumulados}")
        else:
            info = self.info_acumulados_actual
            origenes = self.acumulados["ORIGEN"].value_counts()
            generado = f"generado el {info['generado']:%d/%m/%Y %H:%M}, " if info.get("generado") else ""
            self.info_acumulados.config(
                text=f"ACUMULADOS {generado}del {self.fecha_desde:%d/%m/%Y} al {self.fecha:%d/%m/%Y}: "
                     f"{len(self.acumulados)} leads ({origenes.get('CONTACTACION', 0)} de "
                     f"CONTACTACION, {origenes.get('IA', 0)} de IA)\n{acum.ruta_acumulados()}")
            if (self.selector_desde.get(), self.selector_hasta.get()) != (self.fecha_desde, self.fecha):
                avisos.append("Las fechas elegidas no son las del ACUMULADOS generado; los reportes "
                              "usan las del ACUMULADOS. Vuelve a generar los acumulados para usar las nuevas.")
            if self.historial is not None:
                faltan = acumulados_sin_historial(self.historial, self.acumulados, self.descartados)
                if faltan:
                    avisos.append(f"{len(faltan)} leads de ACUMULADOS no están en el historial: "
                                  "vuelve a descargar el historial.")
        self.aviso_fechas.config(text="\n".join(avisos))

    def _recargar_acumulados(self):
        """Vuelve a leer ACUMULADOS (p. ej. al terminar de generarlos)."""
        self._fijar_acumulados(*cargar_acumulados())
        self._actualizar_nombres()
        self._actualizar_conteo()
        self._actualizar_estado()
        return self.acumulados is not None

    def _al_cambiar_fechas(self):
        self.sel_desde, self.sel_hasta = self.selector_desde.get(), self.selector_hasta.get()
        fn.guardar_preferencias({"fecha_desde": self.sel_desde.isoformat(),
                                 "fecha_hasta": self.sel_hasta.isoformat()})
        if self.acumulados is None:
            self._fijar_acumulados(None, None, self.aviso_acumulados)
            self._actualizar_nombres()
        self._actualizar_estado()

    # --- 1-2. generar acumulados ----------------------------------------------
    def _generar_acumulados(self):
        desde, hasta = self.selector_desde.get(), self.selector_hasta.get()
        if desde > hasta:
            messagebox.showwarning("Periodo inválido", "FECHA_DESDE no puede ser posterior a FECHA_HASTA.")
            return
        try:
            import descargar_base_acumulados as base
            import generar_acumulados as generador
        except BaseException as e:  # noqa: BLE001
            messagebox.showerror("No se pudieron generar los acumulados", str(e))
            return
        self._generando_acumulados = True
        self._actualizar_estado()
        self.config(cursor="watch")
        self.mensajes.delete("1.0", "end")
        self._escribir(f"Generando acumulados del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}...\n")
        self.etapa_acumulados.config(text="Descargando la base de SharePoint...")
        self._cola_acum = queue.Queue()

        def trabajo():
            escritor = _EscritorCola(self._cola_acum)
            try:
                with contextlib.redirect_stdout(escritor), contextlib.redirect_stderr(escritor):
                    base.descargar_base()
                    self._cola_acum.put(("etapa", "Generando ACUMULADOS..."))
                    ruta = generador.generar(desde, hasta)
                self._cola_acum.put(("fin", ruta))
            except BaseException as e:  # noqa: BLE001 - se informa en la ventana
                self._cola_acum.put(("error", e))

        threading.Thread(target=trabajo, daemon=True).start()
        self.after(100, self._revisar_acumulados)

    def _revisar_acumulados(self):
        resultado = None
        while True:
            try:
                mensaje = self._cola_acum.get_nowait()
            except queue.Empty:
                break
            if mensaje[0] == "texto":
                self._escribir(mensaje[1])
            elif mensaje[0] == "etapa":
                self.etapa_acumulados.config(text=mensaje[1])
            else:
                resultado = mensaje
        if resultado is None:
            self.after(100, self._revisar_acumulados)
            return
        self._generando_acumulados = False
        self.config(cursor="")
        if resultado[0] == "fin":
            self.etapa_acumulados.config(text="")
            self._recargar_acumulados()
            if self.acumulados is not None:
                messagebox.showinfo("Acumulados generados",
                                    f"ACUMULADOS del {self.fecha_desde:%d/%m/%Y} al {self.fecha:%d/%m/%Y}: "
                                    f"{len(self.acumulados)} leads.\n\nAhora descarga el historial (paso 4).")
            else:
                messagebox.showerror("No se pudieron generar los acumulados", str(self.aviso_acumulados))
        else:
            error = resultado[1]
            texto = str(error) or type(error).__name__
            self.etapa_acumulados.config(text="No se generaron los acumulados.")
            self._escribir(f"\nERROR: {texto}\n")
            self._actualizar_estado()
            messagebox.showerror("No se pudieron generar los acumulados", texto)

    def _elegir_carpeta(self):
        inicial = self.carpeta_reportes or str(Path.home())
        elegida = filedialog.askdirectory(parent=self, initialdir=inicial, mustexist=True,
                                          title="Carpeta donde se guardarán los reportes")
        if not elegida:
            return
        if not carpeta_escribible(elegida):
            messagebox.showerror("Carpeta sin permisos",
                                 f"No se pueden guardar archivos en:\n{elegida}\n\nElige otra carpeta.")
            return
        self.carpeta_reportes = str(Path(elegida))
        self.texto_carpeta.set(self.carpeta_reportes)
        fn.guardar_preferencias({"carpeta_reportes": self.carpeta_reportes})
        self._actualizar_estado()

    def _abrir_carpeta(self):
        if not self.carpeta_reportes:
            return
        destino = Path(self.carpeta_reportes) / nombre_carpeta_reportes(self.fecha)
        try:
            abrir_en_explorador(destino if destino.is_dir() else self.carpeta_reportes)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("No se pudo abrir la carpeta", str(e))

    def _actualizar_info_historial(self):
        """Muestra la fecha de creación del historial cargado y dónde está."""
        ruta = self.ruta_historial
        self.aviso_historial.config(text="")
        if self.historial is None:
            self.info_historial.config(text=f"Aún no hay historial. Se descargará en:\n{ruta}")
            return
        info = leer_info_historial(ruta)
        try:
            creado = info["descargado"] if info else datetime.fromtimestamp(os.path.getmtime(ruta))
        except OSError:
            creado = None
        fecha = f"{creado:%d/%m/%Y %H:%M}" if creado else "desconocida"
        self.info_historial.config(
            text=f"Historial creado el {fecha}: {self.historial['LEAD_ID'].nunique()} leads, "
                 f"{len(self.historial)} movimientos.\n{ruta}")

    def _avisar_estatus_no_listados(self):
        if self.historial is None:
            return
        no_listados = estatus_no_listados(self.historial, self.estatus)
        if no_listados:
            self._escribir("Aviso: estos estatus del historial no están en la lista de "
                           f"secret.json y sus leads no se pueden elegir: {', '.join(no_listados)}\n")

    # --- descarga del historial -----------------------------------------------
    def _carpeta_disponible(self):
        """Avisa y devuelve False si no hay carpeta o si ya no existe (p. ej. USB desconectada)."""
        if not self.carpeta_reportes:
            messagebox.showwarning("Falta la carpeta", "Elige primero la carpeta de reportes.")
            return False
        if not Path(self.carpeta_reportes).is_dir():
            messagebox.showerror("La carpeta no existe",
                                 f"No se encontró la carpeta de reportes:\n{self.carpeta_reportes}\n\n"
                                 "Conéctala de nuevo o elige otra.")
            return False
        return True

    def _descargar(self):
        if self.acumulados is None:
            messagebox.showwarning("Faltan los acumulados", "Genera primero los acumulados (pasos 1-2).")
            return
        try:
            import extract_data_from_kommo as extractor
        except BaseException as e:  # noqa: BLE001 - p. ej. falta algo en secret.json
            messagebox.showerror("No se pudo preparar la descarga", str(e))
            return

        self._descargando = True
        self._actualizar_estado()
        self.mensajes.delete("1.0", "end")
        self._escribir(f"Descargando el historial de los {len(self.acumulados)} leads de ACUMULADOS...\n")
        self._objetivo = 0.0
        self._mostrado = 0.0
        self._resultado = None
        self._pintar_barra()
        self.etapa_descarga.config(text="Preparando...")
        self._cola = queue.Queue()
        # Destino: configuration.json -> CARPETA_HISTORIAL
        destino = self._ruta_descarga = self.ruta_historial
        self._escribir(f"Se guardará en: {destino}\n")

        def al_avanzar(fraccion, texto):
            self._cola.put(("avance", fraccion, texto))

        def trabajo():
            escritor = _EscritorCola(self._cola, destino)
            try:
                with contextlib.redirect_stdout(escritor), contextlib.redirect_stderr(escritor):
                    extractor.descargar_historial(destino, al_avanzar=al_avanzar)
                anomalias = dict(getattr(extractor, "ULTIMAS_ANOMALIAS", {}) or {})
                self._cola.put(("fin", anomalias))
            except BaseException as e:  # noqa: BLE001 - se informa en la ventana
                self._cola.put(("error", e))

        self._hilo = threading.Thread(target=trabajo, daemon=True)
        self._hilo.start()
        self.after(100, self._revisar_descarga)

    def _pintar_barra(self):
        self.barra["value"] = self._mostrado * 100
        self.porcentaje.config(text=f"{int(self._mostrado * 100)} %")

    def _revisar_descarga(self):
        while True:
            try:
                mensaje = self._cola.get_nowait()
            except queue.Empty:
                break
            tipo = mensaje[0]
            if tipo == "texto":
                self._escribir(mensaje[1])
            elif tipo == "avance":
                _, fraccion, texto = mensaje
                if fraccion is not None:
                    self._objetivo = max(self._objetivo, min(float(fraccion), 1.0))
                if texto:
                    self.etapa_descarga.config(text=texto)
            else:
                self._resultado = mensaje

        # La barra avanza suave hacia el último aviso y, mientras llega el
        # siguiente, sigue moviéndose despacio para que no parezca congelada.
        if self._mostrado < self._objetivo:
            self._mostrado += max(0.004, (self._objetivo - self._mostrado) * 0.2)
        elif self._mostrado < min(self._objetivo + 0.08, 0.97):
            self._mostrado += 0.0015
        self._mostrado = min(self._mostrado, 1.0 if self._resultado else 0.99)
        self._pintar_barra()

        if self._resultado is None:
            self.after(100, self._revisar_descarga)
        else:
            self._terminar_descarga(self._resultado)

    def _terminar_descarga(self, resultado):
        self._descargando = False
        if resultado[0] == "fin":
            _, anomalias = resultado
            historial, aviso = cargar_historial(self._ruta_descarga)
            if historial is None:
                self._fallo_descarga(aviso or "No se pudo leer el historial descargado.")
            else:
                self.historial = historial
                self.etiquetas = cargar_etiquetas(fn.carpeta_anomalos())
                self.descartados = cargar_descartados(fn.carpeta_anomalos())
                guardar_info_historial(self._ruta_descarga, historial)
                self._mostrado = 1.0
                self._pintar_barra()
                self.etapa_descarga.config(text="Descarga terminada.")
                self._actualizar_info_historial()
                self._actualizar_conteo()
                self._avisar_estatus_no_listados()
                self._actualizar_estado()
                texto = (f"Se descargaron {historial['LEAD_ID'].nunique()} leads "
                         f"({len(historial)} movimientos).")
                con_anomalias = {lid for grupo in anomalias.values() for lid in grupo}
                if con_anomalias:
                    texto += (f"\n\n{len(con_anomalias)} leads requieren revisión (varios asesores, "
                              "incompletos o no encontrados); los incompletos no se usan en los "
                              "reportes. El detalle está en el cuadro de mensajes y en "
                              f"{fn.carpeta_anomalos() / 'leads_anomalos.txt'}.")
                messagebox.showinfo("Historial descargado", texto)
                return
        else:
            self._fallo_descarga(explicar_error_descarga(resultado[1]))

    def _fallo_descarga(self, texto):
        self._descargando = False
        self._mostrado = self._objetivo = 0.0
        self._pintar_barra()
        self.porcentaje.config(text="")
        self.etapa_descarga.config(text="La descarga no se completó; se conserva el historial anterior."
                                   if self.historial is not None else "La descarga no se completó.")
        self._escribir(f"\nERROR: {texto}\n")
        self._actualizar_estado()
        messagebox.showerror("No se pudo descargar el historial", texto)

    def _cerrar(self):
        if self._ocupado() and not messagebox.askyesno(
                "Descarga en curso",
                "Se está descargando el historial. Si cierras ahora, la descarga se cancela "
                "y se conserva el historial anterior.\n\n¿Cerrar de todos modos?"):
            return
        self.destroy()

    # --- acciones -------------------------------------------------------------
    def _marcar(self, todos):
        if todos:
            self.lista.select_set(0, "end")
        else:
            self.lista.selection_clear(0, "end")
        self._actualizar_conteo()

    def _estatus_elegidos(self):
        return [self.lista.get(i) for i in self.lista.curselection()]

    def _corte(self):
        """FECHA_HASTA 23:59:59 en modo "hasta el corte"; None en el modo por defecto."""
        return fin_del_dia(self.fecha) if self.modo_eventos.get() == MODO_CORTE else None

    def _al_cambiar_modo(self):
        self._actualizar_conteo()

    def _leads_a_usar(self, estatus):
        """Leads de ACUMULADOS con los estatus elegidos, según el modo de eventos:
        en modo "hasta el corte", sin los movimientos posteriores ni los leads que
        al corte no existían o no tenían asesor."""
        df = filtrar_historial(universo(self.historial, self.acumulados,
                                        self.secret.get("asesores")), estatus)
        corte = self._corte()
        if corte is not None and df is not None:
            df = aplicar_corte(df, corte)
            sin_asesor = {lid for lid, asesores in permitidos_al_corte(self.etiquetas, corte).items()
                          if not asesores}
            df = df[~df["LEAD_ID"].isin(sin_asesor)]
        return df

    def _actualizar_conteo(self):
        if self.historial is None or self.acumulados is None:
            self.conteo.config(text="Leads que cumplen los filtros: —")
            return
        df = self._leads_a_usar(self._estatus_elegidos())
        self.conteo.config(text=f"Leads que cumplen los filtros: {df['LEAD_ID'].nunique()}")

    def _prefijo_valido(self):
        """El prefijo escrito, si es válido; si no, None."""
        try:
            return validar_prefijo(self.prefijo.get())
        except ValueError:
            return None

    def _al_cambiar_prefijo(self):
        self._actualizar_nombres()
        self._actualizar_estado()

    def _actualizar_nombres(self):
        """Muestra dónde y con qué nombre quedarán los reportes del paquete."""
        if hasattr(self, "opcion_corte"):
            self.opcion_corte.config(text=f"Solo hasta el corte: {self.fecha:%d/%m/%Y} a las 23:59:59")
        try:
            prefijo = validar_prefijo(self.prefijo.get())
        except ValueError as e:
            self.vista_prefijo.config(text=str(e), foreground="red")
            return
        fecha = nombre_carpeta_reportes(self.fecha)
        self.vista_prefijo.config(
            text=f"{fecha}/GENERALES/{prefijo}{NOMBRE_ARCHIVO_GENERAL}_{fecha}.xlsx\n"
                 f"{fecha}/INDIVIDUALES/{prefijo}/{prefijo}{PREFIJO_INDIVIDUAL}_<asesor>.xlsx",
            foreground="gray40")

    def _escribir(self, texto):
        _Consola(self).write(texto)

    def _generar(self):
        if not self._carpeta_disponible():
            return
        if self.historial is None:
            messagebox.showwarning("Sin historial", "Descarga primero el historial.")
            return
        if not self._recargar_acumulados():          # siempre con la versión más reciente del archivo
            messagebox.showerror("ACUMULADOS", f"No se pudo leer ACUMULADOS:\n{self.aviso_acumulados}")
            return
        estatus = self._estatus_elegidos()
        if not estatus:
            messagebox.showwarning("Sin estatus", "Elige al menos un estatus de negocio.")
            return
        self.boton.state(["disabled"])
        self.config(cursor="watch")
        self.mensajes.delete("1.0", "end")
        try:
            with contextlib.redirect_stdout(_Consola(self)):
                df = self._leads_a_usar(estatus)
                corte = self._corte()
                print("Eventos: " + (f"hasta el corte, {corte:%d/%m/%Y %H:%M:%S}" if corte else
                                     "hasta el momento de ejecución"))
                print(f"ACUMULADOS: {len(self.acumulados)} leads, del {self.fecha_desde:%d/%m/%Y} "
                      f"al {self.fecha:%d/%m/%Y}")
                faltan = acumulados_sin_historial(self.historial, self.acumulados, self.descartados)
                if faltan:
                    print(f"Aviso: {len(faltan)} leads de ACUMULADOS no están en el historial: "
                          f"{', '.join(map(str, faltan))}")
                print(f"Estatus: {', '.join(estatus)}")
                print(f"Leads seleccionados: {df['LEAD_ID'].nunique()}\n")
                ruta = generar_reportes(df, self.secret, self.fecha, self.descripcion.get().strip(),
                                        self.prefijo.get(), carpeta_reportes=self.carpeta_reportes,
                                        corte=corte, etiquetas=self.etiquetas)
            prefijo = validar_prefijo(self.prefijo.get())
            messagebox.showinfo("Listo", f"Reportes generados en:\n{ruta.parent.parent}\n\n"
                                         f"General: GENERALES/{ruta.name}\n"
                                         f"Individuales: INDIVIDUALES/{prefijo}/")
        except PermissionError as e:
            messagebox.showerror("Archivo abierto",
                                 f"No se pudo reemplazar un archivo. Ciérralo en Excel e intenta de nuevo.\n\n{e}")
        except Exception as e:  # noqa: BLE001 - se muestra cualquier error en pantalla
            self._escribir(f"\nERROR: {e}\n")
            messagebox.showerror("Error", str(e))
        finally:
            self.config(cursor="")
            self._actualizar_estado()


def main():
    try:
        datos = cargar_datos()
    except Exception as e:  # noqa: BLE001
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror("No se pudo iniciar", str(e))
        raiz.destroy()
        return
    App(**datos).mainloop()


if __name__ == "__main__":
    main()