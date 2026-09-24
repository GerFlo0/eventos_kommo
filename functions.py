import json
import locale
import os
from pathlib import Path

import pandas as pd

# Carpeta raíz del proyecto (donde vive este archivo). Sirve para que los
# scripts funcionen aunque se ejecuten desde otra carpeta.
PROJECT_ROOT = Path(__file__).resolve().parent


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
    """Ruta relativa a la carpeta actual, para imprimirla en consola."""
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
