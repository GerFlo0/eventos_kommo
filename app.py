"""
Programa para descargar el historial de etapas de Kommo y generar los
reportes individuales y el reporte general.

Qué leads entran en los reportes: todos los del archivo ACUMULADOS
(configuration.json -> RUTA_ACUMULADOS; hojas "ACUMULADO CONTACTACION" e
"ACUMULADO IA"). El periodo también sale de ahí: su fecha más antigua es
FECHA_DESDE y la más reciente FECHA_HASTA (nombre de la carpeta, título y
"Ventas del día").

En la ventana:
  - se ve el ACUMULADOS cargado (periodo, leads y cuántos no están en el
    historial) y se puede recargar;
  - se elige la carpeta donde se guardan los reportes: todos quedan en
    <carpeta>/<FECHA_HASTA como AAAA-MM-DD>/;
  - se descarga el historial de Kommo (con barra de avance): el historial
    completo de los leads de los asesores, sin límite de fechas. Se guarda en
    configuration.json -> CARPETA_HISTORIAL (se reemplaza en cada descarga),
    junto con su .info.json y el reporte de leads anómalos (.txt y .json).
    La ventana muestra la fecha de creación del historial cargado;
  - se eligen los estatus de negocio a incluir (lista de secret.json);
  - un texto opcional al inicio del nombre de los archivos.

Para generar reportes hacen falta el historial, ACUMULADOS y una carpeta de
reportes elegida. La carpeta se recuerda entre usos (preferencias.json en la
carpeta de datos del programa; ver functions.carpeta_datos()).
"""
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
from general_report import (CARPETA_INDIVIDUALES, NOMBRE_ARCHIVO_GENERAL, PREFIJO_INDIVIDUAL,
                            carpeta_del_dia, generar_reporte_general)
from individual_reports import (_nombre_archivo, aplicar_asignacion, cargar_asignaciones,
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


def validar_prefijo(texto):
    """Texto para el inicio de un nombre de archivo.

    Se quitan solo los espacios del inicio: un espacio al final se respeta,
    porque puede servir de separador ("ENERO " -> "ENERO reporte_general...").
    """
    texto = (texto or "").lstrip()
    prohibidos = sorted(set(texto) & CARACTERES_PROHIBIDOS)
    if prohibidos:
        raise ValueError(f"Caracteres no permitidos en nombres de archivo: {' '.join(prohibidos)}")
    return texto


def nombre_carpeta_reportes(fecha):
    """Subcarpeta de los reportes de un periodo: FECHA_HASTA como AAAA-MM-DD."""
    return f"{fecha:%Y-%m-%d}"


def ruta_historial():
    """Ubicación del historial: configuration.json -> CARPETA_HISTORIAL.
    Todo lo que lee, escribe o describe el historial usa esta función."""
    return fn.ruta_historial()


def generar_reportes(df, secret, fecha, descripcion="", prefijo_individual="", prefijo_general="",
                     carpeta_reportes=None, asignaciones=None):
    """Reportes individuales de cada asesor + reporte general.

    fecha            : FECHA_HASTA; nombra la carpeta y los archivos.
    carpeta_reportes : carpeta elegida en la ventana. Todos los reportes
                       (individuales y general) quedan en
                       <carpeta_reportes>/<AAAA-MM-DD>/. Si no se indica, se usan
                       las carpetas de siempre dentro del proyecto
                       (tablas/reportes/individuales|generales/<mes>/<día>).
    Los prefijos se agregan al inicio del nombre de los reportes individuales
    y del general (vacío = nombre normal).
    asignaciones     : {LEAD_ID: asesor} de los leads con varios asesores (de
                       leads_anomalos.json); cada uno sale solo en el reporte
                       de su asesor. Si no se indica, se lee junto al historial.
    Devuelve la ruta del reporte general.
    """
    prefijo_individual = validar_prefijo(prefijo_individual)
    prefijo_general = validar_prefijo(prefijo_general)

    if carpeta_reportes:
        carpeta_dia = Path(carpeta_reportes) / nombre_carpeta_reportes(fecha)
        carpeta_general = carpeta_dia
    else:
        carpeta_dia = carpeta_del_dia(fn.ruta_proyecto(CARPETA_INDIVIDUALES), fecha)
        carpeta_general = None           # generar_reporte_general usa su carpeta de siempre

    # El reporte general junta los reportes individuales de la carpeta del día
    # que tienen este prefijo: se borran los de una corrida anterior con el
    # mismo prefijo para no mezclar filtros (los de otros prefijos se quedan).
    if carpeta_dia.is_dir():
        for viejo in carpeta_dia.glob(f"{glob.escape(prefijo_individual)}{PREFIJO_INDIVIDUAL}*.xlsx"):
            viejo.unlink()   # PermissionError si está abierto en Excel

    if asignaciones is None:
        asignaciones = cargar_asignaciones(ruta_historial().parent)
    fecha_corte = datetime.now().replace(microsecond=0)  # misma para todos los reportes
    generados = []
    con = db.connect()
    con.register("df", df)   # la consulta de secret.json lee la tabla "df"
    try:
        for asesor in secret["asesores"]:
            resultado = con.execute(secret["query"], [f"%{asesor}%"]).df()
            resultado = aplicar_asignacion(resultado, asesor, asignaciones)
            ruta = generar_reporte_estado_leads(
                resultado, asesor=asesor, fecha_corte=fecha_corte,
                ruta_salida=carpeta_dia / _nombre_archivo(asesor, prefijo_individual))
            if ruta:
                generados.append(ruta)
    finally:
        con.close()

    if not generados:
        raise ValueError("Ningún asesor tiene leads con los filtros elegidos; "
                         "no se generó el reporte general.")
    return generar_reporte_general(fecha=fecha, carpeta_individuales=carpeta_dia,
                                   carpeta_salida=carpeta_general,
                                   descripcion=descripcion, prefijo=prefijo_general,
                                   prefijo_individuales=prefijo_individual)


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


def universo(historial, acumulados):
    """Historial de los leads de ACUMULADOS (todos), con su ORIGEN
    (CONTACTACION / IA). Los demás leads del historial no entran en los reportes."""
    if historial is None or acumulados is None:
        return None
    origen = acumulados.set_index("LEAD_ID")["ORIGEN"]
    df = historial[historial["LEAD_ID"].isin(origen.index)].copy()
    df["ORIGEN"] = df["LEAD_ID"].map(origen)
    return df


def acumulados_sin_historial(historial, acumulados):
    """LEAD_IDs de ACUMULADOS que no están en el historial."""
    if acumulados is None:
        return []
    en_historial = set(historial["LEAD_ID"]) if historial is not None else set()
    return sorted(set(acumulados["LEAD_ID"]) - en_historial)


def cargar_acumulados():
    """(DataFrame, desde, hasta, aviso). Si no se puede leer: (None, hoy, hoy, motivo)."""
    try:
        datos = acum.leer_acumulados()
        desde, hasta = acum.periodo(datos)
        return datos, desde, hasta, None
    except Exception as e:  # noqa: BLE001 - se muestra en la ventana
        hoy = date.today()
        return None, hoy, hoy, str(e)


def cargar_datos():
    """Todo lo que la ventana necesita. Lanza una excepción solo si falta
    algo sin lo que no puede funcionar (secret.json); el historial es opcional
    porque se puede descargar desde la ventana."""
    secret = fn.import_json("json/secret.json")
    estatus = leer_lista_estatus(secret)
    prefs = fn.leer_preferencias()
    acumulados, fecha_desde, fecha_hasta, aviso_acum = cargar_acumulados()   # el periodo sale de ahí
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
        "aviso_acumulados": aviso_acum,
    }


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
                 aviso_inicial=None, acumulados=None, aviso_acumulados=None):
        super().__init__()
        self.acumulados, self.aviso_acumulados = acumulados, aviso_acumulados
        self.secret, self.estatus, self.historial = secret, estatus, historial
        self.fecha, self.fecha_desde = fecha, fecha_desde
        self.carpeta_reportes = carpeta_reportes
        self._descargando = False
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

        # ACUMULADOS (define los leads y el periodo) y carpeta de reportes
        caja = ttk.LabelFrame(izquierda, text="ACUMULADOS y carpeta de reportes", padding=8)
        caja.pack(fill="x")
        self.info_acumulados = ttk.Label(caja, wraplength=440, justify="left")
        self.info_acumulados.pack(anchor="w")
        self.aviso_fechas = ttk.Label(caja, foreground="red", wraplength=440, justify="left")
        self.aviso_fechas.pack(anchor="w")
        self.boton_acumulados = ttk.Button(caja, text="Recargar ACUMULADOS",
                                           command=self._recargar_acumulados)
        self.boton_acumulados.pack(anchor="w", pady=(4, 0))

        ttk.Label(caja, text="Carpeta de reportes:").pack(anchor="w", pady=(4, 0))
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
        caja = ttk.LabelFrame(izquierda, text="Historial de etapas (Kommo)", padding=8)
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

        # Estatus de negocio (selección múltiple)
        caja = ttk.LabelFrame(izquierda, text="Estatus de negocio a incluir", padding=8)
        caja.pack(fill="both", expand=True, pady=(10, 0))
        lista_marco = ttk.Frame(caja)
        lista_marco.pack(fill="both", expand=True)
        self.lista = tk.Listbox(lista_marco, selectmode="multiple", exportselection=False,
                                height=min(max(len(self.estatus), 4), 8))
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
        ttk.Label(caja, text="Descripción para el título del reporte general (opcional):").pack(anchor="w")
        self.descripcion = tk.StringVar()
        ttk.Entry(caja, textvariable=self.descripcion).pack(fill="x")

        # Texto al inicio de los nombres de archivo (vacío = nombre normal)
        caja = ttk.LabelFrame(derecha, text="Texto al inicio del nombre de los archivos (opcional)",
                              padding=8)
        caja.pack(fill="x", pady=(10, 0))
        caja.columnconfigure(1, weight=1)
        self.prefijo_individual = tk.StringVar()
        self.prefijo_general = tk.StringVar()
        self.vista_individual = ttk.Label(caja, foreground="gray40", wraplength=440)
        self.vista_general = ttk.Label(caja, foreground="gray40", wraplength=440)
        for i, (texto, variable, vista) in enumerate([
                ("Reportes individuales:", self.prefijo_individual, self.vista_individual),
                ("Reporte general:", self.prefijo_general, self.vista_general)]):
            ttk.Label(caja, text=texto).grid(row=i * 2, column=0, sticky="w", padx=(0, 8))
            ttk.Entry(caja, textvariable=variable, width=10).grid(row=i * 2, column=1, sticky="ew")
            vista.grid(row=i * 2 + 1, column=0, columnspan=2, sticky="w", pady=(0, 4))
            variable.trace_add("write", lambda *_a: self._actualizar_nombres())
        self._actualizar_nombres()

        # Conteo y botón
        fila = ttk.Frame(derecha)
        fila.pack(fill="x", pady=(10, 0))
        self.conteo = ttk.Label(fila, text="")
        self.conteo.pack(side="left")
        self.boton = ttk.Button(fila, text="Generar reportes", command=self._generar)
        self.boton.pack(side="right")
        self.requisito = ttk.Label(derecha, foreground="red", wraplength=440, justify="left")
        self.requisito.pack(anchor="w", pady=(4, 0))

        self.mensajes = ScrolledText(marco, height=8, wrap="word")
        self.mensajes.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

    # --- estado general -------------------------------------------------------
    def _actualizar_estado(self):
        """Habilita o deshabilita los controles según lo que falte."""
        sin_carpeta = not self.carpeta_reportes
        sin_acumulados = self.acumulados is None
        # La descarga no depende del periodo ni de la carpeta de reportes: el
        # historial va a CARPETA_HISTORIAL (configuration.json).
        puede_descargar = not self._descargando
        puede_generar = (puede_descargar and not sin_carpeta and not sin_acumulados
                         and self.historial is not None)
        self.boton_descargar.state(["!disabled"] if puede_descargar else ["disabled"])
        self.boton.state(["!disabled"] if puede_generar else ["disabled"])
        self.boton_carpeta.state(["disabled"] if self._descargando else ["!disabled"])
        self.boton_abrir.state(["disabled"] if (sin_carpeta or self._descargando) else ["!disabled"])
        self.boton_acumulados.state(["disabled"] if self._descargando else ["!disabled"])
        self._actualizar_info_acumulados()
        if self.historial is None and not self._descargando:
            requisito = "Descarga el historial para poder generar reportes."
        elif sin_carpeta:
            requisito = "Elige la carpeta de reportes para poder generar reportes."
        elif sin_acumulados:
            requisito = "Revisa el archivo de ACUMULADOS (RUTA_ACUMULADOS en configuration.json)."
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

    def _actualizar_info_acumulados(self):
        if self.acumulados is None:
            self.info_acumulados.config(text=f"ACUMULADOS: {acum.ruta_acumulados() or '(sin configurar)'}")
            self.aviso_fechas.config(text=f"No se pudo leer ACUMULADOS: {self.aviso_acumulados}")
            return
        origenes = self.acumulados["ORIGEN"].value_counts()
        faltan = acumulados_sin_historial(self.historial, self.acumulados) if self.historial is not None else []
        self.info_acumulados.config(
            text=f"ACUMULADOS: {acum.ruta_acumulados()}\n"
                 f"Periodo: del {self.fecha_desde:%d/%m/%Y} al {self.fecha:%d/%m/%Y} (FECHA_HASTA)\n"
                 f"{len(self.acumulados)} leads ({origenes.get('CONTACTACION', 0)} de CONTACTACION, "
                 f"{origenes.get('IA', 0)} de IA)")
        self.aviso_fechas.config(
            text=f"{len(faltan)} leads de ACUMULADOS no están en el historial (ver leads_anomalos.txt)."
            if faltan else "")

    def _recargar_acumulados(self):
        """Vuelve a leer ACUMULADOS (p. ej. si se actualizó el archivo)."""
        self.acumulados, self.fecha_desde, self.fecha, self.aviso_acumulados = cargar_acumulados()
        self._actualizar_nombres()
        self._actualizar_conteo()
        self._actualizar_estado()
        return self.acumulados is not None

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
        try:
            import extract_data_from_kommo as extractor
        except BaseException as e:  # noqa: BLE001 - p. ej. falta algo en secret.json
            messagebox.showerror("No se pudo preparar la descarga", str(e))
            return

        self._descargando = True
        self._actualizar_estado()
        self.mensajes.delete("1.0", "end")
        self._escribir("Descargando el historial completo de los leads de los asesores...\n")
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
                    texto += (f"\n\n{len(con_anomalias)} leads requieren revisión (más de un "
                              "asesor, o sin ESTATUS DE NEGOCIO). El detalle está en el cuadro de "
                              f"mensajes y en {self._ruta_descarga.parent / 'leads_anomalos.txt'}.")
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
        if self._descargando and not messagebox.askyesno(
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

    def _leads_a_usar(self, estatus):
        """Leads de ACUMULADOS con los estatus elegidos."""
        return filtrar_historial(universo(self.historial, self.acumulados), estatus)

    def _actualizar_conteo(self):
        if self.historial is None or self.acumulados is None:
            self.conteo.config(text="Leads que cumplen los filtros: —")
            return
        df = self._leads_a_usar(self._estatus_elegidos())
        self.conteo.config(text=f"Leads que cumplen los filtros: {df['LEAD_ID'].nunique()}")

    def _actualizar_nombres(self):
        """Muestra cómo quedará el nombre de los archivos."""
        nombres = (
            (self.prefijo_individual, self.vista_individual, f"{PREFIJO_INDIVIDUAL}_<asesor>.xlsx"),
            (self.prefijo_general, self.vista_general,
             f"{NOMBRE_ARCHIVO_GENERAL}_{self.fecha:%Y-%m-%d}.xlsx"),
        )
        for variable, vista, nombre in nombres:
            try:
                vista.config(text=validar_prefijo(variable.get()) + nombre, foreground="gray40")
            except ValueError as e:
                vista.config(text=str(e), foreground="red")

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
                print(f"ACUMULADOS: {len(self.acumulados)} leads, del {self.fecha_desde:%d/%m/%Y} "
                      f"al {self.fecha:%d/%m/%Y}")
                faltan = acumulados_sin_historial(self.historial, self.acumulados)
                if faltan:
                    print(f"Aviso: {len(faltan)} leads de ACUMULADOS no están en el historial: "
                          f"{', '.join(map(str, faltan))}")
                print(f"Estatus: {', '.join(estatus)}")
                print(f"Leads seleccionados: {df['LEAD_ID'].nunique()}\n")
                ruta = generar_reportes(df, self.secret, self.fecha, self.descripcion.get().strip(),
                                        self.prefijo_individual.get(), self.prefijo_general.get(),
                                        carpeta_reportes=self.carpeta_reportes)
            messagebox.showinfo("Listo", f"Reportes generados en:\n{ruta.parent}\n\n"
                                         f"Reporte general:\n{ruta.name}")
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
