"""
Interfaz para generar los reportes individuales y el reporte general.

Antes hay que correr extract_data_from_kommo.py (genera
tablas/historial_etapas_kommo.xlsx).

En la ventana se elige:
  - qué estatus de negocio incluir (lista de secret.json);
  - si usar todos los leads del historial o solo los creados a partir de
    settings.FECHA_DESDE de configuration.json.

Las carpetas y archivos de los reportes se nombran con settings.FECHA_HASTA
de configuration.json (la fecha de hoy si es null).
"""
import contextlib
import glob
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import duckdb as db
import pandas as pd

import functions as fn
from general_report import (CARPETA_INDIVIDUALES, PREFIJO_INDIVIDUAL,
                            carpeta_del_dia, generar_reporte_general)
from individual_reports import generar_reporte_estado_leads, normalizar

RUTA_HISTORIAL = "tablas/historial_etapas_kommo.xlsx"
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
                         "Vuelve a correr extract_data_from_kommo.py.")


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


def generar_reportes(df, secret, fecha, descripcion="", prefijo_individual="", prefijo_general=""):
    """Reportes individuales de cada asesor + reporte general.

    fecha: nombra las carpetas y archivos. Los prefijos se agregan al inicio
    del nombre de los reportes individuales y del general (vacío = nombre
    normal). Devuelve la ruta del reporte general.
    """
    prefijo_individual = validar_prefijo(prefijo_individual)
    prefijo_general = validar_prefijo(prefijo_general)

    carpeta_base = fn.ruta_proyecto(CARPETA_INDIVIDUALES)
    carpeta_dia = carpeta_del_dia(carpeta_base, fecha)

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
                carpeta_salida=carpeta_base, fecha_archivo=fecha,
                prefijo=prefijo_individual)
            if ruta:
                generados.append(ruta)
    finally:
        con.close()

    if not generados:
        raise ValueError("Ningún asesor tiene leads con los filtros elegidos; "
                         "no se generó el reporte general.")
    return generar_reporte_general(fecha=fecha, carpeta_individuales=carpeta_dia,
                                   descripcion=descripcion, prefijo=prefijo_general,
                                   prefijo_individuales=prefijo_individual)


def cargar_datos():
    """Todo lo que la ventana necesita; lanza una excepción si algo falta."""
    secret = fn.import_json("json/secret.json")
    historial = fn.import_xlsx(RUTA_HISTORIAL)
    validar_historial(historial)
    return {
        "secret": secret,
        "estatus": leer_lista_estatus(secret),
        "historial": historial,
        "fecha": fn.fecha_reportes(),
        "fecha_desde": fn.fecha_desde(),
    }


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


class App(tk.Tk):
    def __init__(self, secret, estatus, historial, fecha, fecha_desde):
        super().__init__()
        self.secret, self.estatus, self.historial = secret, estatus, historial
        self.fecha, self.fecha_desde = fecha, fecha_desde
        self.title("Reportes de leads - Kommo")
        self.minsize(600, 720)
        self._armar()
        self._actualizar_conteo()
        no_listados = estatus_no_listados(historial, estatus)
        if no_listados:
            self._escribir("Aviso: estos estatus del historial no están en la lista de "
                           f"secret.json y sus leads no se pueden elegir: {', '.join(no_listados)}\n")

    # --- construcción ---------------------------------------------------------
    def _armar(self):
        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)

        ttk.Label(marco, text=f"Fecha de los reportes (FECHA_HASTA): {self.fecha:%d/%m/%Y}   |   "
                              f"Leads en el historial: {self.historial['LEAD_ID'].nunique()}"
                  ).pack(anchor="w")

        # Estatus de negocio (selección múltiple)
        caja = ttk.LabelFrame(marco, text="Estatus de negocio a incluir", padding=8)
        caja.pack(fill="both", expand=True, pady=(10, 0))
        lista_marco = ttk.Frame(caja)
        lista_marco.pack(fill="both", expand=True)
        self.lista = tk.Listbox(lista_marco, selectmode="multiple", exportselection=False,
                                height=min(max(len(self.estatus), 4), 12))
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
        caja = ttk.LabelFrame(marco, text="Leads a usar", padding=8)
        caja.pack(fill="x", pady=(10, 0))
        self.solo_desde = tk.BooleanVar(value=False)
        ttk.Radiobutton(caja, text="Todos los leads de historial_etapas_kommo",
                        variable=self.solo_desde, value=False,
                        command=self._actualizar_conteo).pack(anchor="w")
        texto = (f"Solo leads creados a partir de FECHA_DESDE ({self.fecha_desde:%d/%m/%Y})"
                 if self.fecha_desde else "Solo leads creados a partir de FECHA_DESDE (no está definida)")
        opcion = ttk.Radiobutton(caja, text=texto, variable=self.solo_desde, value=True,
                                 command=self._actualizar_conteo)
        opcion.pack(anchor="w")
        if not self.fecha_desde:
            opcion.state(["disabled"])

        # Descripción del reporte general
        caja = ttk.Frame(marco)
        caja.pack(fill="x", pady=(10, 0))
        ttk.Label(caja, text="Descripción para el título del reporte general (opcional):").pack(anchor="w")
        self.descripcion = tk.StringVar()
        ttk.Entry(caja, textvariable=self.descripcion).pack(fill="x")

        # Texto al inicio de los nombres de archivo (vacío = nombre normal)
        caja = ttk.LabelFrame(marco, text="Texto al inicio del nombre de los archivos (opcional)",
                              padding=8)
        caja.pack(fill="x", pady=(10, 0))
        caja.columnconfigure(1, weight=1)
        self.prefijo_individual = tk.StringVar()
        self.prefijo_general = tk.StringVar()
        self.vista_individual = ttk.Label(caja, foreground="gray40", wraplength=420)
        self.vista_general = ttk.Label(caja, foreground="gray40", wraplength=420)
        for i, (texto, variable, vista) in enumerate([
                ("Reportes individuales:", self.prefijo_individual, self.vista_individual),
                ("Reporte general:", self.prefijo_general, self.vista_general)]):
            ttk.Label(caja, text=texto).grid(row=i * 2, column=0, sticky="w", padx=(0, 8))
            ttk.Entry(caja, textvariable=variable).grid(row=i * 2, column=1, sticky="ew")
            vista.grid(row=i * 2 + 1, column=1, sticky="w", pady=(0, 4))
            variable.trace_add("write", lambda *_a: self._actualizar_nombres())
        self._actualizar_nombres()

        # Conteo y botón
        fila = ttk.Frame(marco)
        fila.pack(fill="x", pady=(10, 0))
        self.conteo = ttk.Label(fila, text="")
        self.conteo.pack(side="left")
        self.boton = ttk.Button(fila, text="Generar reportes", command=self._generar)
        self.boton.pack(side="right")

        self.mensajes = ScrolledText(marco, height=10, wrap="word")
        self.mensajes.pack(fill="both", expand=True, pady=(10, 0))

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
        df = filtrar_historial(self.historial, self._estatus_elegidos(), self._fecha_filtro())
        self.conteo.config(text=f"Leads que cumplen los filtros: {df['LEAD_ID'].nunique()}")

    def _actualizar_nombres(self):
        """Muestra cómo quedará el nombre de los archivos."""
        nombres = (
            (self.prefijo_individual, self.vista_individual, f"{PREFIJO_INDIVIDUAL}_<asesor>.xlsx"),
            (self.prefijo_general, self.vista_general, f"reporte_general_{self.fecha:%Y-%m-%d}.xlsx"),
        )
        for variable, vista, nombre in nombres:
            try:
                vista.config(text=validar_prefijo(variable.get()) + nombre, foreground="gray40")
            except ValueError as e:
                vista.config(text=str(e), foreground="red")

    def _escribir(self, texto):
        _Consola(self).write(texto)

    def _generar(self):
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
                                        self.prefijo_individual.get(), self.prefijo_general.get())
            messagebox.showinfo("Listo", f"Reporte general generado:\n{fn.ruta_para_mostrar(ruta)}")
        except PermissionError as e:
            messagebox.showerror("Archivo abierto",
                                 f"No se pudo reemplazar un archivo. Ciérralo en Excel e intenta de nuevo.\n\n{e}")
        except Exception as e:  # noqa: BLE001 - se muestra cualquier error en pantalla
            self._escribir(f"\nERROR: {e}\n")
            messagebox.showerror("Error", str(e))
        finally:
            self.boton.state(["!disabled"])
            self.config(cursor="")


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