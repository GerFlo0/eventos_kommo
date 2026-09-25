# Descargador de eventos de Kommo

Descarga los eventos de cambio de etapa de los leads de Kommo en un rango de fechas, genera un Excel con el historial y el tiempo que cada lead estuvo en cada etapa, y a partir de él arma los reportes individuales por asesor y un reporte general de efectividad. Todo se hace desde una ventana (`app.py`), que también se puede compilar como ejecutable para que otras personas generen reportes sin instalar Python.

## Requisitos previos

* Python 3.14 o superior (es la versión que exige `setup_invironment.py`).
* pip (gestor de paquetes de Python).
* tkinter para la ventana. Viene incluido con Python en Windows y macOS; en Linux puede requerir instalar el paquete `python3-tk`.

## Instalación y configuración

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/GerFlo0/eventos_kommo.git
   cd eventos_kommo
   ```

2. Crear el entorno virtual, instalar dependencias, crear las carpetas necesarias y la plantilla de `json/secret.json`:
   ```bash
   python setup_invironment.py
   ```
   Al terminar, el script indica cómo activar el entorno virtual.

3. Completar `json/secret.json` (no se sube al repositorio) con el token y subdominio de Kommo, los IDs de los embudos, la lista de `asesores`, la consulta `query` y la lista `estatus_negocio`. El token también puede darse con la variable de entorno `KOMMO_TOKEN`, que tiene prioridad sobre el de `secret.json`. La consulta `query` se ejecuta sobre la tabla `df` (el historial) y recibe como `$1` el nombre del asesor entre comodines (`%nombre%`).

4. Ajustar `json/configuration.json` si hace falta (embudos, etapas, campos de la tarjeta, etc.). Los comentarios al inicio de `extract_data_from_kommo.py` explican cada opción. `FECHA_DESDE` y `FECHA_HASTA` solo se usan como valores iniciales de la ventana y al correr el extractor por consola.

## Uso

```bash
python app.py
```

En la ventana:

1. **Periodo:** elegir `FECHA_DESDE` y `FECHA_HASTA` con el calendario.
2. **Carpeta de reportes:** elegir dónde se guardarán. Todos los reportes (individuales y general) quedan en `<carpeta>/<FECHA_HASTA como AAAA-MM-DD>/`. Sin carpeta elegida no se puede descargar ni generar.
3. **Descargar historial:** baja de Kommo el historial del periodo, con barra de avance. Cada descarga reemplaza por completo a la anterior; si la descarga falla, se conserva el historial anterior.
4. **Filtros y nombres:** elegir los estatus de negocio, si se usan todos los leads o solo los creados a partir de `FECHA_DESDE`, la descripción del título y, opcionalmente, un texto al inicio del nombre de cada archivo.
5. **Generar reportes.**

Las fechas y la carpeta elegidas se recuerdan para la próxima vez. Si se cambia el periodo después de descargar, la ventana avisa que el historial corresponde a otro periodo.

Los scripts también se pueden usar por consola, como antes: `python extract_data_from_kommo.py` descarga con las fechas de `configuration.json`, y `general_report.py` acepta `--fecha`, `--descripcion`, `--prefijo` y `--prefijo-individuales` (ver `python general_report.py --help`).

## Dónde guarda el programa sus archivos

El historial (`historial_etapas_kommo.xlsx`) y las preferencias (`preferencias.json`) son archivos del programa; el usuario no elige su ubicación:

* Desde el código fuente: dentro del proyecto (`tablas/historial_etapas_kommo.xlsx` y `preferencias.json`).
* Como ejecutable: en la carpeta de datos del usuario, `%LOCALAPPDATA%\ReportesKommo` en Windows. No se guardan dentro del ejecutable porque en modo "un archivo" se desempaqueta en una carpeta temporal que se borra al cerrar, y en "Archivos de programa" Windows no permite escribir.

## Compilar el ejecutable

Con el entorno virtual activado y `json/secret.json` completo:

```bash
pip install pyinstaller
python compilar.py            # un solo archivo: dist/ReportesKommo.exe
python compilar.py --carpeta  # una carpeta: dist/ReportesKommo/ (abre más rápido)
```

Hay que compilar en el mismo sistema operativo donde se va a usar (para un `.exe` de Windows, compilar en Windows). El ejecutable de un solo archivo tarda unos segundos en abrir porque se desempaqueta cada vez.

**Importante:** el ejecutable lleva dentro `secret.json`, incluido el token de Kommo, y se puede extraer con herramientas comunes. Compártelo solo con personas de confianza y usa un token que puedas revocar si el archivo circula de más.

## Archivos generados

* Historial: hojas HISTORIAL (un renglón por cambio de etapa, con `creacion_de_lead`) y PIVOTE (un renglón por lead, con la primera fecha en cada etapa).
* `<carpeta>/<AAAA-MM-DD>/reporte_individual_dictaminados_<asesor>.xlsx`: un reporte por asesor.
* `<carpeta>/<AAAA-MM-DD>/reporte_general_dictaminados_<AAAA-MM-DD>.xlsx`: reporte general, con las hojas Efectividad y Detalle completo.

Si se escribió un texto de inicio en la ventana, va antes del nombre (p. ej. `ENERO_reporte_general_dictaminados_2026-09-23.xlsx`).

## Pruebas

Las pruebas no se conectan a Kommo (usan una API simulada) ni necesitan `secret.json`:

```bash
python -m unittest discover tests -v
```

Las pruebas de la ventana se omiten automáticamente si no hay pantalla disponible.
