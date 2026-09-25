import json
import locale
import os
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# Carpeta raíz del proyecto (donde vive este archivo). Sirve para que los
# scripts funcionen aunque se ejecuten desde otra carpeta. En el ejecutable
# de PyInstaller es la carpeta donde se desempaquetan sus archivos
# (sys._MEIPASS): ahí están json/configuration.json y json/secret.json.
PROJECT_ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))

# ---------------------------------------------------------------------------
# Ejecutable (PyInstaller)
# ---------------------------------------------------------------------------
# True cuando el programa corre como ejecutable compilado con PyInstaller.
ES_EJECUTABLE = bool(getattr(sys, "frozen", False))
# Nombre de la carpeta de datos del programa cuando corre como ejecutable.
NOMBRE_APP = "ReportesKommo"
# Archivos que el programa GENERA y necesita conservar entre usos.
RUTA_HISTORIAL = "tablas/historial_etapas_kommo.xlsx"
ARCHIVO_PREFERENCIAS = "preferencias.json"


def carpeta_datos() -> Path:
    """Carpeta donde el programa guarda sus propios archivos (historial y
    preferencias). El usuario no la elige.

    - Desde el código fuente: la carpeta del proyecto (tablas/..., como siempre).
    - Como ejecutable: la carpeta de datos del usuario
      (%LOCALAPPDATA%\\ReportesKommo en Windows). No se usa la carpeta del
      programa porque en modo "un archivo" es temporal (se borra al cerrar)
      y en "Archivos de programa" Windows no permite escribir.
    """
    if not ES_EJECUTABLE:
        return PROJECT_ROOT
    base = (os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_DATA_HOME")
            or str(Path.home() / ".local" / "share"))
    carpeta = Path(base) / NOMBRE_APP
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def ruta_historial() -> Path:
    """Ubicación del historial de etapas que descarga y usa app.py."""
    return carpeta_datos() / RUTA_HISTORIAL


def _escribir_json_seguro(ruta, datos):
    """Escribe un JSON sin dejarlo a medias si algo falla (archivo temporal + reemplazo)."""
    ruta = Path(ruta)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    temporal = ruta.with_name(ruta.name + ".tmp")
    temporal.write_text(json.dumps(datos, indent=4, ensure_ascii=False, default=str),
                        encoding="utf-8")
    os.replace(temporal, ruta)


def leer_preferencias() -> dict:
    """Preferencias guardadas por app.py (fechas, carpeta de reportes).

    Si el archivo no existe o está dañado, devuelve {} (se usan los valores
    por defecto) en lugar de impedir que el programa abra.
    """
    ruta = carpeta_datos() / ARCHIVO_PREFERENCIAS
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
        return datos if isinstance(datos, dict) else {}
    except (OSError, ValueError):
        return {}


def guardar_preferencias(cambios: dict) -> dict:
    """Actualiza las preferencias con `cambios` y devuelve el resultado."""
    prefs = leer_preferencias()
    prefs.update(cambios)
    _escribir_json_seguro(carpeta_datos() / ARCHIVO_PREFERENCIAS, prefs)
    return prefs


def ruta_proyecto(ruta) -> Path:
    """Convierte una ruta relativa en ruta dentro de la raíz del proyecto.

    Las rutas absolutas se devuelven sin cambios. Se usa para las rutas
    POR DEFECTO del proyecto (json/, tablas/...), no para las que el
    usuario pasa explícitamente.
    """
    ruta = Path(ruta)
    return ruta if ruta.is_absolute() else PROJECT_ROOT / ruta


def resolver_lectura(ruta) -> Path:
    """Ubica un archivo que se va a LEER.

    Primero busca la ruta tal cual (relativa a la carpeta desde donde se
    ejecuta, como siempre); si no existe y es relativa, la busca dentro de
    la raíz del proyecto.
    """
    ruta = Path(ruta)
    if ruta.exists() or ruta.is_absolute():
        return ruta
    candidato = PROJECT_ROOT / ruta
    return candidato if candidato.exists() else ruta


def ruta_para_mostrar(ruta) -> str:
    """Ruta relativa a la carpeta actual, para imprimirla en consola.

    En el ejecutable se muestra la ruta completa: la "carpeta actual" es
    desde donde se abrió el programa y una ruta relativa no le dice nada
    al usuario.
    """
    if ES_EJECUTABLE:
        return str(Path(ruta).resolve())
    try:
        return os.path.relpath(ruta)
    except ValueError:          # Windows: otra unidad de disco
        return str(ruta)


def import_xlsx(file_path: str, sheet: int = 0) -> pd.DataFrame:
    """
    Imports data from an Excel file and returns a pandas DataFrame.

    Parameters:
    file_path (str): The path to the Excel file.
    sheet (int): The index of the sheet to import.
    if the sheet index is not provided, the first sheet will be imported.

    Returns:
    pd.DataFrame: A DataFrame containing the imported data.

    Raises:
    FileNotFoundError: if the file does not exist (with a clear message).
    """
    ruta = resolver_lectura(file_path)
    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de Excel: {ruta}. "
            "¿Ya corriste extract_data_from_kommo.py?")
    return pd.read_excel(ruta, sheet_name=sheet)


def import_json(file_path: str) -> dict:
    """
    Imports data from a JSON file and returns a dictionary.

    Parameters:
    file_path (str): The path to the JSON file.

    Returns:
    dict: A dictionary containing the imported data.

    Raises:
    FileNotFoundError: if the file does not exist.
    ValueError: if the file is not valid JSON (shows line and column).
    """
    ruta = resolver_lectura(file_path)
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo JSON: {ruta}")
    try:
        try:
            # UTF-8 (con o sin BOM) es lo que guardan VS Code y la mayoría
            # de los editores.
            texto = ruta.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            # Compatibilidad: archivo guardado con la codificación del
            # sistema (p. ej. ANSI/cp1252 en Windows), como se leía antes.
            try:
                texto = ruta.read_text(encoding=locale.getpreferredencoding(False))
            except UnicodeDecodeError as e:
                raise ValueError(
                    f"No se pudo leer {ruta}: guárdalo con codificación UTF-8.") from e
        return json.loads(texto)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"El archivo {ruta} no es un JSON válido "
            f"(línea {e.lineno}, columna {e.colno}): {e.msg}") from e

# ---------------------------------------------------------------------------
# Fechas de configuration.json
# ---------------------------------------------------------------------------
RUTA_CONFIGURACION = "json/configuration.json"

def _fecha_de_configuracion(clave):
    """settings.<clave> de configuration.json como date (None si es null)."""
    valor = import_json(RUTA_CONFIGURACION)["settings"].get(clave)
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(f"{clave} en configuration.json debe tener formato "
                         f"AAAA-MM-DD (valor actual: {valor!r})") from None

def fecha_reportes() -> date:
    """Fecha con la que se nombran las carpetas y archivos de los reportes.
    Es settings.FECHA_HASTA de configuration.json; si es null, hoy.
    """
    return _fecha_de_configuracion("FECHA_HASTA") or date.today()

def fecha_desde():
    """settings.FECHA_DESDE de configuration.json como date (None si es null)."""
    return _fecha_de_configuracion("FECHA_DESDE")
