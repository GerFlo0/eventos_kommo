#!/usr/bin/env python3
"""
setup_invironment.py

Prepara el entorno de trabajo del proyecto:
  1. Crea el entorno virtual (si no existe).
  2. Actualiza pip e instala las dependencias de requirements.txt.
  3. Crea las carpetas que el proyecto necesita pero no están en el repositorio.

Uso:
    python setup_invironment.py              # configuración normal
    python setup_invironment.py --recreate   # borra y vuelve a crear el entorno virtual
    python setup_invironment.py --skip-install
"""

import argparse
import os
import shutil
import subprocess
import sys
import venv
from pathlib import Path

# ---------------------------------------------------------------------------
# CONFIGURACIÓN: ajusta estos valores a tu proyecto
# ---------------------------------------------------------------------------

# Nombre de la carpeta del entorno virtual
VENV_DIR_NAME = ".venv"

# Archivo de dependencias
REQUIREMENTS_FILE = "requirements.txt"

# Carpetas a crear (rutas relativas a la raíz del proyecto).
# Se admiten subcarpetas, p. ej. "data/raw".
REQUIRED_FOLDERS = [
    "tablas/reportes/individuales",
    "tablas/reportes/generales"
]

# Versión mínima de Python requerida
MIN_PYTHON = (3, 14)

# ---------------------------------------------------------------------------

# Raíz del proyecto = carpeta donde está este script,
# así funciona sin importar desde dónde se ejecute.
PROJECT_ROOT = Path(__file__).resolve().parent
VENV_DIR = PROJECT_ROOT / VENV_DIR_NAME


def log(msg: str) -> None:
    print(f"[setup] {msg}")


def error(msg: str) -> None:
    print(f"[setup] ERROR: {msg}", file=sys.stderr)


def check_python_version() -> None:
    if sys.version_info < MIN_PYTHON:
        error(
            f"Se requiere Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} o superior. "
            f"Versión actual: {sys.version.split()[0]}"
        )
        sys.exit(1)


def get_venv_python() -> Path:
    """Ruta al ejecutable de Python dentro del venv (Windows vs Linux/macOS)."""
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def create_venv(recreate: bool) -> None:
    if VENV_DIR.exists() and recreate:
        log(f"Eliminando entorno virtual existente en {VENV_DIR} ...")
        shutil.rmtree(VENV_DIR)

    if get_venv_python().exists():
        log(f"El entorno virtual ya existe en {VENV_DIR}")
        return

    log(f"Creando entorno virtual en {VENV_DIR} ...")
    venv.EnvBuilder(with_pip=True, clear=False).create(VENV_DIR)
    log("Entorno virtual creado.")


def run(cmd: list) -> None:
    """Ejecuta un comando y aborta si falla."""
    try:
        subprocess.run(cmd, check=True, cwd=PROJECT_ROOT)
    except subprocess.CalledProcessError as exc:
        error(f"Falló el comando: {' '.join(map(str, cmd))} (código {exc.returncode})")
        sys.exit(exc.returncode)


def install_requirements() -> None:
    python = str(get_venv_python())
    req_path = PROJECT_ROOT / REQUIREMENTS_FILE

    log("Actualizando pip ...")
    run([python, "-m", "pip", "install", "--upgrade", "pip"])

    if not req_path.exists():
        log(f"No se encontró {REQUIREMENTS_FILE}; se omite la instalación de dependencias.")
        return

    log(f"Instalando dependencias desde {REQUIREMENTS_FILE} ...")
    run([python, "-m", "pip", "install", "-r", str(req_path)])
    log("Dependencias instaladas.")


def create_folders() -> None:
    log("Creando carpetas necesarias ...")
    for folder in REQUIRED_FOLDERS:
        path = PROJECT_ROOT / folder
        if path.exists():
            log(f"  ya existe: {folder}")
        else:
            path.mkdir(parents=True, exist_ok=True)
            log(f"  creada:    {folder}")


def print_activation_hint() -> None:
    if os.name == "nt":
        activate = f"{VENV_DIR_NAME}\\Scripts\\activate"
    else:
        activate = f"source {VENV_DIR_NAME}/bin/activate"
    print()
    log("Entorno listo. Para activarlo ejecuta:")
    print(f"    {activate}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Configura el entorno del proyecto.")
    parser.add_argument(
        "--recreate",
        action="store_true",
        help="Elimina el entorno virtual existente y lo crea de nuevo.",
    )
    parser.add_argument(
        "--skip-install",
        action="store_true",
        help="No instala dependencias (solo crea venv y carpetas).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    check_python_version()

    log(f"Raíz del proyecto: {PROJECT_ROOT}")
    create_venv(recreate=args.recreate)

    if not args.skip_install:
        install_requirements()

    create_folders()
    print_activation_hint()


if __name__ == "__main__":
    main()