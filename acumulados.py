"""
Lectura del archivo ACUMULADOS (configuration.json -> RUTA_ACUMULADOS).

Define qué leads entran en los reportes: todos los de sus dos hojas.
  - "ACUMULADO CONTACTACION": columnas FECHA y LEAD
  - "ACUMULADO IA":           columnas fecha y ID Lead

Las columnas se buscan por su nombre (sin importar mayúsculas, acentos ni
espacios de más), no por su letra, por si alguien inserta una columna.
El periodo de los reportes sale de las fechas del archivo: la más antigua es
FECHA_DESDE y la más reciente FECHA_HASTA.
"""
import unicodedata
from pathlib import Path

import pandas as pd

import functions as fn

# hoja -> (columna de fecha, columna de lead, ORIGEN que se muestra en los reportes)
HOJAS = {
    "ACUMULADO CONTACTACION": ("FECHA", "LEAD", "CONTACTACION"),
    "ACUMULADO IA": ("fecha", "ID Lead", "IA"),
}


def _normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto or ""))
    return " ".join("".join(c for c in texto if unicodedata.category(c) != "Mn").lower().split())


def ruta_acumulados():
    """configuration.json -> RUTA_ACUMULADOS. Relativa: dentro de la carpeta de
    datos del programa. None si no está configurada."""
    valor = fn.import_json(fn.RUTA_CONFIGURACION)["settings"].get("RUTA_ACUMULADOS")
    if not valor:
        return None
    ruta = Path(str(valor)).expanduser()
    return ruta if ruta.is_absolute() else fn.carpeta_datos() / ruta


def _columna(df, nombre, hoja, ruta):
    buscado = _normalizar(nombre)
    for col in df.columns:
        if _normalizar(col) == buscado:
            return col
    raise ValueError(f"La hoja '{hoja}' de {ruta} no tiene la columna '{nombre}'.")


def leer_acumulados(ruta=None):
    """DataFrame con LEAD_ID, FECHA y ORIGEN (una fila por lead).

    Si un lead aparece más de una vez, se toma su fecha más antigua.
    Lanza FileNotFoundError / ValueError con un mensaje claro si algo falta.
    """
    ruta = Path(ruta) if ruta else ruta_acumulados()
    if ruta is None:
        raise ValueError("Falta configurar RUTA_ACUMULADOS en configuration.json.")
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo de ACUMULADOS: {ruta}")
    libro = pd.read_excel(ruta, sheet_name=None)
    hojas = {_normalizar(h): h for h in libro}
    partes = []
    for hoja, (col_fecha, col_lead, origen) in HOJAS.items():
        real = hojas.get(_normalizar(hoja))
        if real is None:
            raise ValueError(f"El archivo {ruta} no tiene la hoja '{hoja}'.")
        df = libro[real]
        parte = pd.DataFrame({
            "LEAD_ID": pd.to_numeric(df[_columna(df, col_lead, hoja, ruta)], errors="coerce"),
            "FECHA": pd.to_datetime(df[_columna(df, col_fecha, hoja, ruta)], errors="coerce"),
            "ORIGEN": origen,
        })
        partes.append(parte.dropna(subset=["LEAD_ID"]))
    datos = pd.concat(partes, ignore_index=True)
    datos["LEAD_ID"] = datos["LEAD_ID"].astype("int64")
    datos = datos.sort_values(["LEAD_ID", "FECHA"], na_position="last").drop_duplicates("LEAD_ID")
    return datos.reset_index(drop=True)


def periodo(acumulados):
    """(FECHA_DESDE, FECHA_HASTA) como date: la fecha más antigua y la más reciente."""
    fechas = acumulados["FECHA"].dropna()
    if fechas.empty:
        raise ValueError("ACUMULADOS no tiene fechas válidas.")
    return fechas.min().date(), fechas.max().date()
