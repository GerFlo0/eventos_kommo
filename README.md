# Descargador de eventos de Kommo

Descarga los eventos de cambio de etapa de los leads de Kommo en un rango de fechas, genera un Excel con el historial y el tiempo que cada lead estuvo en cada etapa, y a partir de él arma reportes individuales por asesor y un reporte general de efectividad.

## Requisitos previos

* Python 3.14 o superior (es la versión que exige `setup_invironment.py`).
* pip (gestor de paquetes de Python).

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

3. Crear `json/secret.json` (no se sube al repositorio). Estructura esperada, con valores de ejemplo:
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
       "people": {
           "sin_financiera_restringida": ["NOMBRE ASESOR 1", "NOMBRE ASESOR 2"]
       },
       "query": "SELECT * FROM df WHERE LEAD_ID IN (SELECT LEAD_ID FROM df WHERE ETIQUETAS ILIKE $1)"
   }
   ```
   El token también puede darse con la variable de entorno `KOMMO_TOKEN`, que tiene prioridad sobre el de `secret.json`. La consulta `query` se ejecuta sobre la tabla `df` (el historial) y recibe como `$1` el nombre del asesor entre comodines (`%nombre%`).

4. Ajustar `json/configuration.json` según lo que se necesite (rango de fechas, embudos, etapas, campos de la tarjeta, etc.). Los comentarios al inicio de `extract_data_from_kommo.py` explican cada opción.

## Uso

Los scripts se ejecutan en este orden (pueden correrse desde cualquier carpeta):

```bash
python extract_data_from_kommo.py   # 1. descarga de Kommo -> tablas/historial_etapas_kommo.xlsx
python app.py                       # 2. un reporte individual por asesor
python general_report.py            # 3. reporte general de efectividad del día
```

`general_report.py` acepta `--fecha dd/mm/aaaa` para consolidar los reportes de otro día y `--descripcion "texto"` para agregar un texto al título. Ejecuta `python general_report.py --help` para ver todas las opciones.

## Archivos generados

* `tablas/historial_etapas_kommo.xlsx`: hojas HISTORIAL (un renglón por cambio de etapa) y PIVOTE (un renglón por lead, con la primera fecha en cada etapa).
* `tablas/reportes/individuales/<mes>/<día>/reporte_estado_leads_<asesor>.xlsx`: un reporte por asesor.
* `tablas/reportes/generales/<mes>/<día>/reporte_general_<aaaa-mm-dd>.xlsx`: reporte general.

La carpeta `tablas/` y los archivos de Excel están excluidos del repositorio porque contienen datos de leads.

## Pruebas

Las pruebas no se conectan a Kommo (usan una API simulada) ni necesitan `secret.json`:

```bash
python -m unittest discover tests -v
```
