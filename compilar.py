#!/usr/bin/env python3
"""
compilar.py

Genera el ejecutable de app.py con PyInstaller, para que otras personas
puedan descargar el historial y generar reportes sin instalar Python.

Uso (con el entorno virtual activado):
    pip install pyinstaller
    python compilar.py            # un solo archivo: dist/ReportesKommo(.exe)
    python compilar.py --carpeta  # una carpeta: dist/ReportesKommo/ (abre más rápido)

Qué se empaqueta:
  - json/configuration.json y json/secret.json (solo lectura dentro del programa).

Qué NO se empaqueta:
  - El historial y las preferencias: el programa los guarda en la carpeta de
    datos de cada usuario (%LOCALAPPDATA%\\ReportesKommo en Windows).

IMPORTANTE: el ejecutable lleva dentro secret.json, incluido el TOKEN de
Kommo. Cualquiera que tenga el ejecutable puede extraerlo. Compártelo solo
con personas de confianza y usa un token que puedas revocar.

Hay que compilar en el mismo sistema operativo donde se va a usar: para un
.exe de Windows, correr este script en Windows.
"""
import argparse
import os
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
NOMBRE = "ReportesKommo"
DATOS = ["json/configuration.json", "json/secret.json"]


def main():
    parser = argparse.ArgumentParser(description="Compila app.py con PyInstaller.")
    parser.add_argument("--carpeta", action="store_true",
                        help="Genera una carpeta en lugar de un solo archivo (abre más rápido).")
    args = parser.parse_args()

    try:
        import PyInstaller.__main__ as pyinstaller
    except ImportError:
        sys.exit("Falta PyInstaller. Instálalo con:  pip install pyinstaller")

    faltan = [d for d in DATOS if not (RAIZ / d).is_file()]
    if faltan:
        sys.exit(f"Faltan archivos para empaquetar: {', '.join(faltan)}")

    opciones = [
        str(RAIZ / "app.py"),
        "--name", NOMBRE,
        "--noconfirm",
        "--clean",
        "--windowed",                          # sin ventana de consola
        "--onedir" if args.carpeta else "--onefile",
        "--distpath", str(RAIZ / "dist"),
        "--workpath", str(RAIZ / "build"),
        "--specpath", str(RAIZ / "build"),
        # Se importa dentro de una función de app.py; se declara por si acaso.
        "--hidden-import", "extract_data_from_kommo",
    ]
    for dato in DATOS:
        opciones += ["--add-data", f"{RAIZ / dato}{os.pathsep}{Path(dato).parent}"]

    print("Compilando... (puede tardar varios minutos)")
    pyinstaller.run(opciones)
    destino = RAIZ / "dist" / (NOMBRE if args.carpeta else NOMBRE + (".exe" if os.name == "nt" else ""))
    print(f"\nListo: {destino}")
    print("Recuerda: el ejecutable incluye el token de Kommo de secret.json.")


if __name__ == "__main__":
    main()
