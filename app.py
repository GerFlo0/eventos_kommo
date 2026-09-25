"""
Programa para descargar el historial de etapas de Kommo y generar los
reportes individuales y el reporte general.

En la ventana se elige:
  - el periodo (FECHA_DESDE y FECHA_HASTA) con un calendario;
  - la carpeta donde se guardan los reportes: todos quedan en
    <carpeta>/<FECHA_HASTA como AAAA-MM-DD>/;
  - descargar el historial de Kommo (con barra de avance). Se guarda en la
    carpeta de datos del programa y cada descarga reemplaza a la anterior;
  - qué estatus de negocio incluir (lista de secret.json);
  - si usar todos los leads del historial o solo los creados a partir de
    FECHA_DESDE;
  - un texto opcional al inicio del nombre de los archivos.

Sin carpeta de reportes elegida no se puede descargar ni generar.
Las fechas y la carpeta se recuerdan entre usos (preferencias.json en la
carpeta de datos del programa; ver functions.carpeta_datos()).
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

import functions as fn
from general_report import (CARPETA_INDIVIDUALES, NOMBRE_ARCHIVO_GENERAL, PREFIJO_INDIVIDUAL,
                            carpeta_del_dia, generar_reporte_general)
from individual_reports import MESES, _nombre_archivo, generar_reporte_estado_leads, normalizar

COL_ESTATUS = "ESTATUS DE NEGOCIO"
COL_CREACION = "creacion_de_lead"
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
    faltan = [c for c in ("LEAD_ID", COL_ESTATUS, COL_CREACION) if c not in df.columns]
    if faltan:
        raise ValueError(f"Al historial le faltan las columnas {faltan}. "
                         "Descarga el historial de nuevo.")


def filtrar_historial(df, estatus, solo_desde_fecha=None):
    """Historial completo de los leads que cumplen los filtros.

    estatus          : estatus de negocio elegidos (sin importar acentos ni
                       mayúsculas). Un lead entra si alguno de sus estatus
                       está en la lista.
    solo_desde_fecha : date; si se indica, solo entran los leads creados ese
                       día o después. Los leads sin fecha de creación quedan
                       fuera.
    """
    elegidos = {normalizar(e) for e in estatus}
    por_lead = _estatus_por_lead(df)
    leads = set(por_lead[por_lead.map(lambda v: bool(_estatus_de_celda(v) & elegidos))].index)
    if solo_desde_fecha is not None:
        creacion = pd.to_datetime(df.groupby("LEAD_ID")[COL_CREACION].first(), errors="coerce")
        leads &= set(creacion[creacion >= pd.Timestamp(solo_desde_fecha)].index)
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


def generar_reportes(df, secret, fecha, descripcion="", prefijo_individual="", prefijo_general="",
                     carpeta_reportes=None):
    """Reportes individuales de cada asesor + reporte general.

    fecha            : FECHA_HASTA; nombra la carpeta y los archivos.
    carpeta_reportes : carpeta elegida en la ventana. Todos los reportes
                       (individuales y general) quedan en
                       <carpeta_reportes>/<AAAA-MM-DD>/. Si no se indica, se usan
                       las carpetas de siempre dentro del proyecto
                       (tablas/reportes/individuales|generales/<mes>/<día>).
    Los prefijos se agregan al inicio del nombre de los reportes individuales
    y del general (vacío = nombre normal). Devuelve la ruta del reporte general.
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

    fecha_corte = datetime.now().replace(microsecond=0)  # misma para todos los reportes
    generados = []
    con = db.connect()
    con.register("df", df)   # la consulta de secret.json lee la tabla "df"
    try:
        for asesor in secret["asesores"]:
            resultado = con.execute(secret["query"], [f"%{asesor}%"]).df()
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


# ---- Historial de etapas (archivo del programa) ---------------------------
def ruta_info_historial(ruta_historial):
    """Datos de la última descarga, junto al historial (…historial_etapas_kommo.info.json)."""
    ruta = Path(ruta_historial)
    return ruta.with_name(f"{ruta.stem}.info.json")


def guardar_info_historial(ruta_historial, fecha_desde, fecha_hasta, historial):
    """Anota cuándo y con qué periodo se descargó el historial."""
    info = {
        "descargado": datetime.now().isoformat(timespec="seconds"),
        "fecha_desde": fecha_desde.isoformat(),
        "fecha_hasta": fecha_hasta.isoformat(),
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
        info["fecha_desde"] = date.fromisoformat(info["fecha_desde"])
        info["fecha_hasta"] = date.fromisoformat(info["fecha_hasta"])
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
        historial = pd.read_excel(ruta)
        validar_historial(historial)
        return historial, None
    except Exception as e:  # noqa: BLE001 - archivo dañado o de una versión anterior
        return None, f"El historial guardado no se pudo usar ({e}). Descárgalo de nuevo."


def _fecha_de_texto(valor):
    try:
        return date.fromisoformat(valor) if valor else None
    except (TypeError, ValueError):
        return None


def cargar_datos():
    """Todo lo que la ventana necesita. Lanza una excepción solo si falta
    algo sin lo que no puede funcionar (secret.json); el historial es opcional
    porque se puede descargar desde la ventana."""
    secret = fn.import_json("json/secret.json")
    estatus = leer_lista_estatus(secret)
    prefs = fn.leer_preferencias()
    hoy = date.today()
    fecha_desde = (_fecha_de_texto(prefs.get("fecha_desde")) or fn.fecha_desde()
                   or hoy.replace(day=1))
    fecha_hasta = _fecha_de_texto(prefs.get("fecha_hasta")) or fn.fecha_reportes()
    carpeta = prefs.get("carpeta_reportes") or None
    ruta = fn.ruta_historial()
    historial, aviso = cargar_historial(ruta)
    return {
        "secret": secret,
        "estatus": estatus,
        "historial": historial,
        "fecha": fecha_hasta,
        "fecha_desde": fecha_desde,
        "carpeta_reportes": carpeta,
        "ruta_historial": ruta,
        "aviso_inicial": aviso,
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
                 ruta_historial=None, aviso_inicial=None):
        super().__init__()
        self.secret, self.estatus, self.historial = secret, estatus, historial
        self.fecha, self.fecha_desde = fecha, fecha_desde
        self.carpeta_reportes = carpeta_reportes
        self.ruta_historial = Path(ruta_historial or fn.ruta_historial())
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

        # Periodo y carpeta de reportes
        caja = ttk.LabelFrame(izquierda, text="Periodo y carpeta de reportes", padding=8)
        caja.pack(fill="x")
        fechas = ttk.Frame(caja)
        fechas.pack(fill="x")
        ttk.Label(fechas, text="Desde (FECHA_DESDE):").grid(row=0, column=0, sticky="w")
        self.selector_desde = SelectorFecha(fechas, self.fecha_desde, self._al_cambiar_fechas)
        self.selector_desde.grid(row=0, column=1, sticky="w", padx=(6, 0), pady=(0, 3))
        ttk.Label(fechas, text="Hasta (FECHA_HASTA):").grid(row=1, column=0, sticky="w")
        self.selector_hasta = SelectorFecha(fechas, self.fecha, self._al_cambiar_fechas)
        self.selector_hasta.grid(row=1, column=1, sticky="w", padx=(6, 0))
        self.aviso_fechas = ttk.Label(caja, foreground="red")
        self.aviso_fechas.pack(anchor="w")

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

        # Qué leads usar
        caja = ttk.LabelFrame(derecha, text="Leads a usar", padding=8)
        caja.pack(fill="x")
        self.solo_desde = tk.BooleanVar(value=False)
        ttk.Radiobutton(caja, text="Todos los leads de historial_etapas_kommo",
                        variable=self.solo_desde, value=False,
                        command=self._actualizar_conteo).pack(anchor="w")
        self.opcion_desde = ttk.Radiobutton(caja, variable=self.solo_desde, value=True,
                                            command=self._actualizar_conteo)
        self.opcion_desde.pack(anchor="w")
        self._actualizar_texto_desde()

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
        fechas_mal = self.selector_desde.get() > self.selector_hasta.get()
        puede_descargar = not (self._descargando or sin_carpeta or fechas_mal)
        puede_generar = puede_descargar and self.historial is not None
        self.boton_descargar.state(["!disabled"] if puede_descargar else ["disabled"])
        self.boton.state(["!disabled"] if puede_generar else ["disabled"])
        self.boton_carpeta.state(["disabled"] if self._descargando else ["!disabled"])
        self.boton_abrir.state(["disabled"] if (sin_carpeta or self._descargando) else ["!disabled"])
        for selector in (self.selector_desde, self.selector_hasta):
            selector.habilitar(not self._descargando)

        self.aviso_fechas.config(
            text="FECHA_DESDE no puede ser posterior a FECHA_HASTA." if fechas_mal else "")
        if sin_carpeta:
            requisito = "Elige la carpeta de reportes para poder descargar el historial y generar reportes."
        elif fechas_mal:
            requisito = "Corrige el periodo para continuar."
        elif self.historial is None and not self._descargando:
            requisito = "Descarga el historial para poder generar reportes."
        else:
            requisito = ""
        self.requisito.config(text=requisito)
        if self.carpeta_reportes:
            destino = Path(self.carpeta_reportes) / nombre_carpeta_reportes(self.selector_hasta.get())
            self.destino.config(text=f"Los reportes se guardarán en: {destino}")
        else:
            self.destino.config(text="")

    def _al_cambiar_fechas(self):
        self.fecha_desde = self.selector_desde.get()
        self.fecha = self.selector_hasta.get()
        fn.guardar_preferencias({"fecha_desde": self.fecha_desde.isoformat(),
                                 "fecha_hasta": self.fecha.isoformat()})
        self._actualizar_texto_desde()
        self._actualizar_nombres()
        self._actualizar_conteo()
        self._actualizar_info_historial()
        self._actualizar_estado()

    def _actualizar_texto_desde(self):
        self.opcion_desde.config(
            text=f"Solo leads creados a partir de FECHA_DESDE ({self.fecha_desde:%d/%m/%Y})")

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
        if self.historial is None:
            self.info_historial.config(text="Aún no hay historial descargado.")
            self.aviso_historial.config(text="")
            return
        leads = self.historial["LEAD_ID"].nunique()
        info = leer_info_historial(self.ruta_historial)
        if info is None:
            self.info_historial.config(text=f"Historial guardado: {leads} leads, "
                                            f"{len(self.historial)} movimientos (periodo desconocido).")
            self.aviso_historial.config(text="")
            return
        self.info_historial.config(
            text=f"Descargado el {info['descargado']:%d/%m/%Y %H:%M}: periodo "
                 f"{info['fecha_desde']:%d/%m/%Y} – {info['fecha_hasta']:%d/%m/%Y}, "
                 f"{leads} leads, {len(self.historial)} movimientos.")
        distinto = (info["fecha_desde"], info["fecha_hasta"]) != (self.fecha_desde, self.fecha)
        self.aviso_historial.config(
            text="El periodo elegido no coincide con el del historial descargado; "
                 "descárgalo de nuevo si quieres usar el periodo nuevo." if distinto else "")

    def _avisar_estatus_no_listados(self):
        if self.historial is None:
            return
        no_listados = estatus_no_listados(self.historial, self.estatus)
        if no_listados:
            self._escribir("Aviso: estos estatus del historial no están en la lista de "
                           f"secret.json y sus leads no se pueden elegir: {', '.join(no_listados)}\n")

    # --- descarga del historial -----------------------------------------------
    def _descargar(self):
        if not self.carpeta_reportes:
            messagebox.showwarning("Falta la carpeta", "Elige primero la carpeta de reportes.")
            return
        desde, hasta = self.selector_desde.get(), self.selector_hasta.get()
        if desde > hasta:
            messagebox.showwarning("Periodo inválido", "FECHA_DESDE no puede ser posterior a FECHA_HASTA.")
            return
        try:
            import extract_data_from_kommo as extractor
        except BaseException as e:  # noqa: BLE001 - p. ej. falta algo en secret.json
            messagebox.showerror("No se pudo preparar la descarga", str(e))
            return

        self._descargando = True
        self._actualizar_estado()
        self.mensajes.delete("1.0", "end")
        self._escribir(f"Descargando historial del {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}...\n")
        self._objetivo = 0.0
        self._mostrado = 0.0
        self._resultado = None
        self._pintar_barra()
        self.etapa_descarga.config(text="Preparando...")
        self._cola = queue.Queue()

        def al_avanzar(fraccion, texto):
            self._cola.put(("avance", fraccion, texto))

        def trabajo():
            escritor = _EscritorCola(self._cola, self.ruta_historial)
            try:
                with contextlib.redirect_stdout(escritor), contextlib.redirect_stderr(escritor):
                    extractor.descargar_historial(desde.isoformat(), hasta.isoformat(),
                                                  self.ruta_historial, al_avanzar=al_avanzar)
                self._cola.put(("fin", desde, hasta))
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
            _, desde, hasta = resultado
            historial, aviso = cargar_historial(self.ruta_historial)
            if historial is None:
                self._fallo_descarga(aviso or "No se pudo leer el historial descargado.")
            else:
                self.historial = historial
                guardar_info_historial(self.ruta_historial, desde, hasta, historial)
                self._mostrado = 1.0
                self._pintar_barra()
                self.etapa_descarga.config(text="Descarga terminada.")
                self._actualizar_info_historial()
                self._actualizar_conteo()
                self._avisar_estatus_no_listados()
                self._actualizar_estado()
                messagebox.showinfo("Historial descargado",
                                    f"Se descargaron {historial['LEAD_ID'].nunique()} leads "
                                    f"({len(historial)} movimientos).")
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

    def _fecha_filtro(self):
        return self.fecha_desde if self.solo_desde.get() else None

    def _actualizar_conteo(self):
        if self.historial is None:
            self.conteo.config(text="Leads que cumplen los filtros: —")
            return
        df = filtrar_historial(self.historial, self._estatus_elegidos(), self._fecha_filtro())
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
        if not self.carpeta_reportes:
            messagebox.showwarning("Falta la carpeta", "Elige primero la carpeta de reportes.")
            return
        if self.historial is None:
            messagebox.showwarning("Sin historial", "Descarga primero el historial.")
            return
        if self.fecha_desde > self.fecha:
            messagebox.showwarning("Periodo inválido", "FECHA_DESDE no puede ser posterior a FECHA_HASTA.")
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
                df = filtrar_historial(self.historial, estatus, self._fecha_filtro())
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
