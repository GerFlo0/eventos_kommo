import requests
import functions as fn
from pathlib import Path

config = fn.import_json("json/configuration.json")

archivos = config["acumulados"]
carpeta_destino = fn.PROJECT_ROOT / "acumulados"
carpeta_destino.mkdir(parents=True, exist_ok=True)

for nombre, url in archivos.items():
    respuesta = requests.get(
        url,
        params={"download": "1"},
        allow_redirects=True
    )
    respuesta.raise_for_status()

    nombre_archivo = nombre if Path(nombre).suffix else f"{nombre}.xlsx"
    destino = carpeta_destino / nombre_archivo
    destino.write_bytes(respuesta.content)
    print(f"Archivo descargado correctamente: {destino.name}")