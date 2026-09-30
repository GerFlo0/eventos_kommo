"""
Descarga los archivos base de los acumulados desde SharePoint.

Los enlaces están en secret.json -> sharepoint ({nombre: enlace}); cada
archivo se guarda como <CARPETA_BASE_ACUMULADOS>/<nombre>.xlsx
(configuration.json). generar_acumulados.py los usa según la clave "BASE" de
cada fuente.

Si algo falla (sin internet, enlace vencido, SharePoint pide iniciar sesión),
se detiene con un mensaje claro y no deja archivos a medias.

Uso:
    python descargar_base_acumulados.py
"""
import sys
from pathlib import Path

import requests

import functions as fn


class ErrorDescarga(Exception):
    """Error con mensaje listo para mostrar al usuario."""


def ruta_base(nombre, carpeta=None):
    """Archivo local de una base: <carpeta>/<nombre>.xlsx."""
    carpeta = Path(carpeta) if carpeta else fn.carpeta_base_acumulados()
    return carpeta / (nombre if Path(nombre).suffix else f"{nombre}.xlsx")


def enlaces_sharepoint():
    """{nombre: enlace} de secret.json -> sharepoint."""
    enlaces = fn.import_json("json/secret.json").get("sharepoint") or {}
    if not enlaces:
        raise ErrorDescarga("Faltan los enlaces de SharePoint en secret.json -> sharepoint.")
    return enlaces


def descargar_base(carpeta=None, enlaces=None):
    """Descarga cada archivo base. Devuelve las rutas guardadas.

    Cada archivo se escribe primero con otro nombre y solo reemplaza al
    anterior si se descargó completo y es un Excel.
    """
    enlaces = enlaces or enlaces_sharepoint()
    guardados = []
    for nombre, url in enlaces.items():
        destino = ruta_base(nombre, carpeta)
        destino.parent.mkdir(parents=True, exist_ok=True)
        print(f'- Descargando "{nombre}" de SharePoint...')
        try:
            respuesta = requests.get(url, params={"download": "1"}, allow_redirects=True, timeout=120)
        except requests.RequestException as e:
            raise ErrorDescarga(f'No se pudo descargar "{nombre}": revisa tu conexión a internet.'
                                f"\n\nDetalle: {e}") from e
        if respuesta.status_code >= 400:
            raise ErrorDescarga(f'SharePoint respondió con error {respuesta.status_code} al descargar '
                                f'"{nombre}": el enlace pudo vencer o ya no tener permiso.')
        # Un .xlsx es un archivo ZIP: siempre empieza con "PK". Si no, SharePoint
        # devolvió otra cosa (normalmente la página para iniciar sesión).
        if not respuesta.content.startswith(b"PK"):
            raise ErrorDescarga(f'SharePoint no devolvió un archivo de Excel para "{nombre}": el '
                                "enlace pide iniciar sesión o ya no es válido. Revisa el enlace en "
                                "secret.json -> sharepoint.")
        temporal = destino.with_name(destino.name + ".descargando")
        try:
            temporal.write_bytes(respuesta.content)
            temporal.replace(destino)
        except PermissionError as e:
            temporal.unlink(missing_ok=True)
            raise ErrorDescarga(f"No se pudo guardar {destino}. ¿Está abierto en Excel? "
                                "Ciérralo e intenta de nuevo.") from e
        print(f"  Guardado: {destino}")
        guardados.append(destino)
    return guardados


def main():
    try:
        descargar_base()
    except ErrorDescarga as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
