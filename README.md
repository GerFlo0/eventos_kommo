# Reportes de leads Kommo

Descarga de Kommo el historial de cambios de etapa de los leads de los embudos CIERRES y VENTAS, y genera a partir de él un reporte individual por asesor y un reporte general de efectividad. Todo se hace desde una ventana (`app.py`), que también se puede compilar como ejecutable para que otras personas generen reportes sin instalar Python.

## Qué hace

* Descarga los cambios de etapa de **cualquier etapa** de los embudos CIERRES y VENTAS, de los leads que tienen al menos una etiqueta que contiene **"CIERRES"** (sin importar mayúsculas ni acentos).
* De la tarjeta de cada lead toma: ESTATUS DE NEGOCIO, FECHA DICTAMEN, MONTO OTORGADO y MOTIVO LEAD PERDIDO. Si un campo está vacío, la celda queda vacía y el lead no se descarta.
* Revisa los leads descargados y reporta los que tienen anomalías: más de una etiqueta con "CIERRES", sin FECHA DICTAMEN o sin ESTATUS DE NEGOCIO.
* Genera un reporte por cada asesor de la lista de `secret.json` y un reporte general que los consolida.

## Requisitos

* Python 3.14 o superior (es la versión que exige `setup_invironment.py`).
* pip.
* tkinter para la ventana. Viene incluido con Python en Windows y macOS; en Linux puede requerir el paquete `python3-tk`.

## Instalación

1. Clonar el repositorio:
```bash
   git clone https://github.com/GerFlo0/eventos_kommo.git
   cd eventos_kommo
```

2. Crear el entorno virtual, instalar dependencias, crear las carpetas necesarias y la plantilla de `json/secret.json`:
```bash
   python setup_invironment.py
```
   Opciones: `--recreate` borra y vuelve a crear el entorno virtual; `--skip-install` no instala dependencias. Al terminar, el script indica cómo activar el entorno virtual.

3. Completar `json/secret.json` (no se sube al repositorio).

4. Revisar `json/configuration.json`.

## Configuración

### `json/secret.json`

| Clave | Para qué sirve |
|---|---|
| `kommo.TOKEN` | Token de larga duración de Kommo. También puede darse con la variable de entorno `KOMMO_TOKEN`, que tiene prioridad. |
| `kommo.SUBDOMAIN` | Subdominio de la cuenta (`<subdominio>.kommo.com`). |
| `kommo.PIPELINE_ID` | IDs de los embudos `CIERRES` y `VENTAS`. Si `configuration.json` trae el `ID` del embudo, se usa ese. |
| `asesores` | Lista de asesores de los que se genera reporte. Un asesor que no esté en la lista no tendrá reporte individual ni aparecerá en el general. |
| `query` | Consulta que se ejecuta sobre el historial (tabla `df`) para obtener los leads de cada asesor. Recibe como `$1` el nombre del asesor entre comodines (`%nombre%`). |
| `estatus_negocio` | Lista de estatus de negocio que se pueden elegir en la ventana. Si la guardas con otra clave, ajusta `CLAVES_LISTA_ESTATUS` en `app.py`. |

La consulta usa los nombres internos de las columnas (`LEAD_ID`, `ETIQUETAS`, `FECHA`...). Aunque la hoja HISTORIAL del Excel tenga títulos como "LEAD ID" o "FECHA EVENTO", al cargarla los reportes los traducen a los internos, así que la consulta no necesita cambios. Ejemplo:

```sql
SELECT * FROM df WHERE LEAD_ID IN (SELECT LEAD_ID FROM df WHERE ETIQUETAS ILIKE $1)
```

### `json/configuration.json`

| Clave | Para qué sirve |
|---|---|
| `FECHA_DESDE`, `FECHA_HASTA` | Periodo (`AAAA-MM-DD`). En la ventana solo son los valores iniciales la primera vez; por consola definen el periodo. `FECHA_HASTA` se incluye completa; `null` = hoy. |
| `INCLUIR_LEAD_ADDED` | Incluye el evento de creación del lead como primer movimiento. |
| `SALIDA` | Archivo que genera el extractor por consola. La ventana no lo usa. |
| `ARCHIVOS_SEPARADOS` | Por consola: `true` genera un Excel por embudo. La ventana siempre genera un solo archivo. |
| `ETIQUETA_LEADS` | Texto que debe contener alguna etiqueta del lead para incluirlo (`"CIERRES"`). |
| `CAMPOS_TARJETA` | Campos de la tarjeta que se agregan como columnas, en ese orden. Se buscan por nombre en Kommo, sin importar mayúsculas ni acentos. |
| `CAMPOS_REQUERIDOS` | Campos que el lead debe tener llenos para incluirlo. `[]` = no se descarta ningún lead. |
| `CAMPOS_DINERO` | Campos que se guardan como número con formato de dinero (`MONTO OTORGADO`). |
| `EMBUDOS` | Embudos a descargar, con `NOMBRE`, `ID` y, opcionalmente, `CLAVE_SECRET` (clave en `secret.json` si no se indica `ID`). Se toman todas sus etapas. |
| `performance` | `THREADS` (descargas en paralelo), `VERBOSE` (detalle por consola), `INCLUIR_USUARIOS` (agrega al final una columna con quién movió el lead). |

## Uso

```bash
python app.py
```

En la ventana:

1. **Periodo:** elegir `FECHA_DESDE` y `FECHA_HASTA` con el calendario.
2. **Carpeta de reportes:** elegir dónde se guardará todo. Lo de cada periodo queda en `<carpeta>/<FECHA_HASTA como AAAA-MM-DD>/`. Sin carpeta elegida no se puede descargar ni generar.
3. **Descargar historial:** baja el historial del periodo con barra de avance y lo guarda en la carpeta del periodo. Si ya había uno ahí, se reemplaza; los de otros periodos no se tocan. Si la descarga falla, se conserva el que había. Al terminar, muestra los leads con anomalías.
4. **Filtros:** elegir los estatus de negocio y los leads a usar: todos los del historial, o solo los que tienen su FECHA DICTAMEN entre `FECHA_DESDE` y `FECHA_HASTA` (ambos días incluidos). En este último modo, los leads sin FECHA DICTAMEN quedan fuera.
5. **Opcional:** descripción para el título del reporte general y un texto al inicio del nombre de los archivos (por separado para los individuales y el general).
6. **Generar reportes.**

Las fechas y la carpeta elegidas se recuerdan. Al abrir el programa, o al cambiar la carpeta o `FECHA_HASTA`, se usa el historial que ya exista en `<carpeta>/<FECHA_HASTA>/`; si no hay, hay que descargarlo.

### Uso por consola

```bash
python extract_data_from_kommo.py   # descarga con el periodo de configuration.json -> SALIDA
python general_report.py --help     # opciones: --fecha, --descripcion, --prefijo, --prefijo-individuales
```

Por consola, las anomalías se muestran en pantalla pero no se guardan en archivo.

## Archivos generados

En `<carpeta de reportes>/<AAAA-MM-DD>/`:

| Archivo | Contenido |
|---|---|
| `historial_etapas_kommo.xlsx` | Hoja **HISTORIAL**: un renglón por cambio de etapa. Hoja **PIVOTE**: un renglón por lead, con la primera fecha en cada etapa. |
| `historial_etapas_kommo.info.json` | Fecha de descarga y periodo del historial. |
| `leads_anomalos.txt` | Leads con anomalías, agrupados por tipo. Se crea en cada descarga, aunque no haya anomalías. |
| `reporte_individual_dictaminados_<asesor>.xlsx` | Un reporte por asesor: Resumen, Estado actual e Historial por lead. |
| `reporte_general_dictaminados_<AAAA-MM-DD>.xlsx` | Hojas Efectividad y Detalle completo. |

Si se escribió un texto de inicio, va antes del nombre (p. ej. `ENERO_reporte_general_dictaminados_2026-09-30.xlsx`). Al generar de nuevo con el mismo texto de inicio, se reemplazan los reportes anteriores con ese texto; los de otros textos se conservan.

### Columnas de la hoja HISTORIAL

`LEAD ID`, `EMBUDO`, `FECHA CREACION DE LEAD`, `ESTATUS DE NEGOCIO`, `FECHA DICTAMEN`, `MONTO OTORGADO` (formato de dinero), `MOTIVO LEAD PERDIDO`, `ETIQUETAS` (solo en el movimiento más reciente de cada lead), `ETAPA ANTERIOR`, `ETAPA_NUEVA`, `FECHA EVENTO`, `DIAS EN ETAPA ANTERIOR`.

### Clasificación en los reportes

Cada lead se clasifica según su etapa actual. En la categoría **Sin Capacidad** entran los leads en la etapa SIN CAPACIDAD y también los que tienen "SIN CAPACIDAD" en MOTIVO LEAD PERDIDO, aunque su etapa sea otra (por ejemplo, "Venta perdida"). La regla se aplica igual en el Resumen de los reportes individuales y en la tabla de Efectividad del reporte general.

## Dónde guarda el programa sus preferencias

`preferencias.json` (carpeta de reportes y fechas elegidas):

* Desde el código fuente: en la carpeta del proyecto.
* Como ejecutable: en `%LOCALAPPDATA%\ReportesKommo` (Windows).

## Compilar el ejecutable

Con el entorno virtual activado y `json/secret.json` completo, **en Windows** (PyInstaller solo genera programas para el sistema donde se ejecuta):

```bash
pip install pyinstaller
python compilar.py
```

Se genera en modo carpeta:

* `dist/ReportesKommo/`: la carpeta del programa, con `ReportesKommo.exe` y la subcarpeta `_internal/`.
* `dist/ReportesKommo.zip`: la misma carpeta comprimida, para compartir.

Quien lo reciba descomprime el `.zip` y abre `ReportesKommo.exe`. El `.exe` no funciona fuera de su carpeta.

**Importante:** el programa lleva dentro `secret.json`, incluido el token de Kommo (en `_internal/json/`). Compártelo solo con personas de confianza y usa un token que puedas revocar.

## Estructura del proyecto

| Archivo | Función |
|---|---|
| `app.py` | Ventana del programa. |
| `extract_data_from_kommo.py` | Descarga de Kommo y armado del historial. |
| `individual_reports.py` | Reportes individuales y clasificación de etapas. |
| `general_report.py` | Reporte general de efectividad. |
| `functions.py` | Utilidades comunes (lectura de archivos, rutas, fechas, preferencias). |
| `setup_invironment.py` | Preparación del entorno. |
| `compilar.py` | Compilación del ejecutable. |
| `json/configuration.json` | Configuración. |
| `json/secret.json` | Credenciales y listas (no se sube al repositorio). |