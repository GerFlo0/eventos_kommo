# Descargador de eventos de Kommo

Descarga los eventos de cambio de etapa de los leads de Kommo en un rango de fechas, genera un Excel con el historial y el tiempo que cada lead estuvo en cada etapa, y a partir de él arma reportes individuales por asesor y un reporte general de efectividad.

## Requisitos previos

* Python 3.14 o superior (es la versión que exige `setup_invironment.py`).
* pip (gestor de paquetes de Python).
* tkinter para la ventana de `app.py`. Viene incluido con Python en Windows y macOS; en Linux puede requerir instalar el paquete `python3-tk`.

## Instalación y configuración

1. Clonar el repositorio:
```bash
   git clone https://github.com/GerFlo0/eventos_kommo.git
   cd eventos_kommo
```

2. Crear el entorno virtual, instalar dependencias y crear las carpetas necesarias:
```bash
   python setup_invironment.py
```
   Al terminar, el script indica cómo activar el entorno virtual.

3. `setup_invironment.py` crea `json/secret.json` (no se sube al repositorio) con valores genéricos. Reemplaza los valores de `kommo`, la lista de `asesores` y el resto de la configuración con los datos de tu cuenta:
```json
   {
       "kommo": {
           "SUBDOMAIN": "tu_subdominio",
           "TOKEN": "token_de_larga_duracion",
           "PIPELINE_ID": {
               "CIERRES": 1234567,
               "VENTAS": 7654321
           }
       },
         "asesores": ["NOMBRE ASESOR 1", "NOMBRE ASESOR 2"],
       "estatus_negocio": ["ESTATUS 1", "ESTATUS 2"],
       "query": "SELECT * FROM df WHERE LEAD_ID IN (SELECT LEAD_ID FROM df WHERE ETIQUETAS ILIKE $1)"
   }
```
   El token también puede darse con la variable de entorno `KOMMO_TOKEN`, que tiene prioridad sobre el de `secret.json`. La consulta `query` se ejecuta sobre la tabla `df` (el historial) y recibe como `$1` el nombre del asesor entre comodines (`%nombre%`). `estatus_negocio` es la lista de estatus que se pueden elegir en la ventana de `app.py` (si la guardas con otra clave, ajusta `CLAVES_LISTA_ESTATUS` en `app.py`).

4. Ajustar `json/configuration.json` según lo que se necesite (rango de fechas, embudos, etapas, campos de la tarjeta, etc.). Los comentarios al inicio de `extract_data_from_kommo.py` explican cada opción.

## Uso

Los scripts se ejecutan en este orden (pueden correrse desde cualquier carpeta):

```bash
python extract_data_from_kommo.py   # 1. descarga de Kommo -> tablas/historial_etapas_kommo.xlsx
python app.py                       # 2. ventana para generar los reportes individuales y el general
```

En la ventana de `app.py` se eligen los estatus de negocio a incluir y si se usan todos los leads del historial o solo los creados a partir de `FECHA_DESDE`. Antes de generar, muestra cuántos leads cumplen los filtros. Al generar, se borran los reportes individuales anteriores de esa misma fecha para no mezclarlos con los nuevos.

Las carpetas y archivos de los reportes se nombran con `FECHA_HASTA` de `configuration.json` (la fecha de hoy si es `null`). El tiempo transcurrido dentro de los reportes se sigue calculando hasta el momento en que se generan.

`general_report.py` también puede correrse solo: acepta `--fecha dd/mm/aaaa` y `--descripcion "texto"` (ver `python general_report.py --help`).

## Archivos generados

* `tablas/historial_etapas_kommo.xlsx`: hojas HISTORIAL (un renglón por cambio de etapa, con la fecha de creación del lead en `creacion_de_lead`) y PIVOTE (un renglón por lead, con la primera fecha en cada etapa).
* `tablas/reportes/individuales/<mes>/<día>/reporte_estado_leads_<asesor>.xlsx`: un reporte por asesor (mes y día de `FECHA_HASTA`).
* `tablas/reportes/generales/<mes>/<día>/reporte_general_<aaaa-mm-dd>.xlsx`: reporte general.

La carpeta `tablas/` está excluida del repositorio porque contiene datos de leads.

## Pruebas

Las pruebas no se conectan a Kommo (usan una API simulada) ni necesitan `secret.json`:

```bash
python -m unittest discover tests -v
```

Las pruebas de la ventana se omiten automáticamente si no hay pantalla disponible.