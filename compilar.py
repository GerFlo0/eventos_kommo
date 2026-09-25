#!/usr/bin/env python3
"""
compilar.py

Genera el ejecutable de app.py con PyInstaller en modo CARPETA (onedir),
para que otras personas puedan descargar el historial y generar reportes sin
instalar Python.

Uso (con el entorno virtual activado):
    pip install pyinstaller
    python compilar.py

Resultado:
    dist/ReportesKommo/                  <- carpeta del programa
        ReportesKommo.exe                <- lo que se abre (en Windows)
        _internal/                       <- librerías y archivos del programa
    dist/ReportesKommo.zip               <- la misma carpeta comprimida, para compartir

Para repartirlo se comparte el .zip; quien lo reciba lo descomprime y abre
ReportesKommo.exe. El .exe NO funciona si se saca de su carpeta (necesita
_internal junto a él); para tenerlo a mano, crear un acceso directo.

Por qué en carpeta y no en un solo archivo: abre más rápido (no se
desempaqueta cada vez), los antivirus lo marcan menos y los errores son más
fáciles de diagnosticar.

Qué se empaqueta:
  - json/configuration.json y json/secret.json (solo lectura dentro del programa).
Qué NO se empaqueta:
  - El historial: app.py lo descarga en la carpeta de reportes que elija el usuario.
  - Las preferencias: se guardan en %LOCALAPPDATA%\\ReportesKommo de cada usuario.

IMPORTANTE:
  - El ejecutable lleva dentro secret.json, incluido el TOKEN de Kommo, y se
    puede extraer. Compártelo solo con personas de confianza y usa un token
    que puedas revocar.
  - PyInstaller no genera programas para otro sistema operativo: para un .exe
    de Windows, este script se tiene que correr EN WINDOWS.
"""
import os
import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
NOMBRE = "ReportesKommo"
DATOS = ["json/configuration.json", "json/secret.json"]


def main():
    try:
        import PyInstaller.__main__ as pyinstaller
    except ImportError:
        sys.exit("Falta PyInstaller. Instálalo con:  pip install pyinstaller")

    faltan = [d for d in DATOS if not (RAIZ / d).is_file()]
    if faltan:
        sys.exit(f"Faltan archivos para empaquetar: {', '.join(faltan)}")

    dist = RAIZ / "dist"
    opciones = [
        str(RAIZ / "app.py"),
        "--name", NOMBRE,
        "--onedir",                            # una carpeta, no un solo archivo
        "--noconfirm",                         # reemplaza la compilación anterior
        "--clean",
        "--windowed",                          # sin ventana de consola
        "--distpath", str(dist),
        "--workpath", str(RAIZ / "build"),
        "--specpath", str(RAIZ / "build"),
        # Se importa dentro de una función de app.py; se declara por si acaso.
        "--hidden-import", "extract_data_from_kommo",
    ]
    for dato in DATOS:
        opciones += ["--add-data", f"{RAIZ / dato}{os.pathsep}{Path(dato).parent}"]

    print("Compilando... (puede tardar varios minutos)")
    pyinstaller.run(opciones)

    carpeta = dist / NOMBRE
    ejecutable = carpeta / (NOMBRE + (".exe" if os.name == "nt" else ""))
    if not ejecutable.is_file():
        sys.exit(f"No se encontró el ejecutable esperado: {ejecutable}")

    print("Comprimiendo la carpeta para compartirla...")
    archivo_zip = shutil.make_archive(str(dist / NOMBRE), "zip", root_dir=dist, base_dir=NOMBRE)

    print(f"\nListo.\n  Programa: {ejecutable}\n  Para compartir: {archivo_zip}")
    print("Recuerda: el programa incluye el token de Kommo de secret.json.")


if __name__ == "__main__":
    main()
