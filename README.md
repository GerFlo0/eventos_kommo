# Reportes de leads Kommo

Programa para generar los reportes de efectividad de los asesores de cierres a partir de tres fuentes:

1. **Base de acumulados** (SharePoint): con ella se genera el archivo ACUMULADOS del periodo, que define qué leads entran en los reportes.
2. **Kommo**: de cada lead de ACUMULADOS se descarga su tarjeta y su historial de etapas hasta la fecha de corte.
3. **Reglas de clasificación**: con ellas se arman paquetes de reportes, cada uno con un reporte individual por asesor y un reporte general (efectividad, monto otorgado, ventas del día y detalle).

Todo se hace desde una ventana (`app.py`), que también se puede compilar como ejecutable para Windows.

## Requisitos

* Python 3.14 o superior (es la versión que exige `setup_invironment.py`).
* pip.
* tkinter para la ventana. Viene incluido con Python en Windows y macOS; en Linux puede requerir el paquete `python3-tk`.
* Acceso a internet (SharePoint y Kommo).

## Instalación

```bash
git clone https://github.com/GerFlo0/eventos_kommo.git
cd eventos_kommo
python setup_invironment.py
```

`setup_invironment.py` crea el entorno virtual, instala las dependencias y crea la plantilla de `json/secret.json`. Opciones: `--recreate` vuelve a crear el entorno virtual; `--skip-install` no instala dependencias. Al terminar, indica cómo activar el entorno virtual.

La plantilla de `secret.json` trae valores de relleno (por ejemplo, los IDs de los embudos) que hay que reemplazar antes de usar el programa.

## Configuración

### `json/secret.json` (no se sube al repositorio)

| Clave | Para qué sirve |
|---|---|
| `kommo.TOKEN` | Token de larga duración de Kommo. También puede darse con la variable de entorno `KOMMO_TOKEN`, que tiene prioridad. |
| `kommo.SUBDOMAIN` | Subdominio de la cuenta (`<subdominio>.kommo.com`). |
| `kommo.PIPELINE_ID` | IDs de los embudos `VENTAS` y `CIERRES`. Si un ID no existe en Kommo, la descarga se detiene y muestra los embudos disponibles. |
| `asesores` | Asesores de los que se genera reporte. Deben escribirse igual que en sus etiquetas de Kommo, con las mismas mayúsculas y acentos. |
| `query` | Consulta que obtiene los leads de cada asesor sobre el historial (tabla `df`). Recibe como `$1` el nombre del asesor entre comodines (`%nombre%`). Ejemplo: `SELECT * FROM df WHERE LEAD_ID IN (SELECT LEAD_ID FROM df WHERE ETIQUETAS LIKE $1)`. |
| `estatus_negocio` | Estatus de negocio que se pueden elegir en la ventana. |
| `campos_tarjeta` | Campos de la tarjeta del lead que se descargan (ESTATUS DE NEGOCIO, FECHA DICTAMEN, MONTO OTORGADO, MOTIVO LEAD PERDIDO). |
| `campos_dinero` | Campos que se guardan con formato de dinero (MONTO OTORGADO). |
| `sharepoint` | Enlaces de descarga de los archivos base: `{"SOLICITUD DE INFORMES 07": "enlace", "SOLICITUD IA PRUEBA": "enlace"}`. Los nombres deben coincidir con la clave `BASE` de cada fuente en `configuration.json`. |

La consulta usa los nombres internos de las columnas (`LEAD_ID`, `ETIQUETAS`). Al cargar el historial, los reportes traducen a esos nombres los títulos del Excel ("LEAD", "FECHA EVENTO"...).

### `json/configuration.json`

**`settings`**

| Clave | Para qué sirve |
|---|---|
| `FECHA_DESDE`, `FECHA_HASTA` | Fechas iniciales del selector (`DD/MM/AAAA`). Después, la ventana recuerda las últimas usadas. |
| `CARPETA_BASE_ACUMULADOS` | Donde se descargan los archivos base de SharePoint. |
| `RUTA_ACUMULADOS` | Archivo ACUMULADOS que se genera. |
| `CARPETA_HISTORIAL` | Donde se guarda `historial_etapas_kommo.xlsx`. |
| `CARPETA_ANOMALOS` | Donde se guardan `leads_anomalos.txt` y `leads_anomalos.json`. |
| `performance` | `THREADS` (descargas en paralelo) y `VERBOSE` (detalle en consola). |

Las rutas relativas se toman dentro de la carpeta de datos del programa: la del proyecto desde el código fuente, o `%LOCALAPPDATA%\ReportesKommo` en el ejecutable.

**`fuentes`**: una por hoja de ACUMULADOS.

| Clave | Para qué sirve |
|---|---|
| `HOJA_DESTINO` | Hoja que se crea en ACUMULADOS (`ACUMULADO CONTACTACION`, `ACUMULADO IA`). |
| `BASE` | Nombre del archivo base, el mismo de `secret.json -> sharepoint`. |
| `HOJA_ORIGEN` | Hoja del archivo base. |
| `COLUMNA_ESTATUS`, `ESTATUS_INCLUIDOS` | Solo entran las filas con esos estatus. |
| `COLUMNA_LEAD`, `LEADS_EXCLUIDOS` | Leads que se descartan (opcional). |
| `COLUMNA_FECHA` | Fecha con la que se filtra el periodo (ambos días incluidos). |
| `COLUMNA_FECHA_RESPALDO` | Opcional: se usa si `COLUMNA_FECHA` está vacía (en IA, `fecha` cuando aún no hay `FECHA ASIG`, como Cloud). |
| `COLUMNAS_SALIDA` | Columnas que se copian a ACUMULADOS, en ese orden. |

Las columnas se buscan por nombre, sin importar mayúsculas, acentos ni espacios de más.

## Uso

```bash
python app.py
```

1. **Periodo:** elegir FECHA_DESDE y FECHA_HASTA con el calendario.
2. **Generar acumulados:** descarga la base de SharePoint y genera ACUMULADOS con ese periodo. El periodo queda anotado en `ACUMULADOS.info.json`, y los reportes y el historial lo usan (fecha de corte, carpeta, título y ventas del día). Si algo falla (sin internet, enlace vencido o que pide iniciar sesión, archivo abierto), se detiene, lo explica y no deja archivos a medias.
3. **Carpeta de reportes:** donde se guardan los paquetes de reportes (ver "Archivos generados").
4. **Descargar historial:** trae de Kommo la tarjeta y los movimientos (en cualquier embudo, hasta la fecha de corte de ACUMULADOS) de todos y únicamente los leads de ACUMULADOS. Se guardan todos, aunque estén incompletos. La descarga reemplaza al historial anterior solo si termina bien.
5. **(Opcional) Estatus de negocio** a incluir.
6. **(Opcional) Descripción** para el título del reporte general.
7. **Prefijo (obligatorio):** identifica el paquete de reportes (por ejemplo, `sin_restringido_` o `solo_restringido_`). Va al inicio del nombre de cada reporte y nombra la subcarpeta de sus reportes individuales. Se quitan los espacios de los extremos. No puede contener `< > : " / \ | ? *`, terminar en punto ni ser un nombre reservado de Windows (CON, NUL, COM1...). Se recomienda usar `_` como separador.
8. **Generar reportes.** Si ya existía un paquete con el mismo prefijo en esa fecha, se reemplaza; los demás paquetes no se tocan.

### Uso por consola

```bash
python descargar_base_acumulados.py                               # descarga la base de SharePoint
python generar_acumulados.py --desde 01/09/2026 --hasta 29/09/2026
python extract_data_from_kommo.py                                 # historial de los leads de ACUMULADOS
python general_report.py --help                                   # --fecha, --carpeta, --salida, --prefijo, ...
```

## Reglas de los reportes

**Leads que se usan:** los de ACUMULADOS que están completos, es decir, con FECHA DICTAMEN, ESTATUS DE NEGOCIO (de los elegidos) y la etiqueta de un asesor de `secret.json`. Los incompletos se listan en `leads_anomalos.txt`.

**Fecha de corte:** los movimientos de Kommo posteriores a la FECHA_HASTA de ACUMULADOS se descartan, así que la etapa de cada lead es la que tenía al cierre del periodo. Los datos de la tarjeta (estatus, fecha de dictamen, monto, motivo, etiquetas, buzón) son los actuales, porque Kommo no guarda su historia.

**Categorías**, según la etapa del lead (sin importar mayúsculas ni acentos):

| Categoría | Etapas |
|---|---|
| GANADOS | OTORGADO (VENTAS), otorgado (CIERRES), Leads ganados |
| OFERTA / OFERTA EN ESPERA / DOCUMENTACION / CAPTURADO | La etapa del mismo nombre, en cualquiera de los dos embudos |
| LEAD PERDIDO | LEAD PERDIDO / lead Perdido, NO INTERESADO |
| SIN CAPACIDAD | SIN CAPACIDAD, sin capacidad (CARTERA), cualquier etapa no clasificada (p. ej. Contactado), y los leads perdidos cuyo MOTIVO LEAD PERDIDO incluye SIN CAPACIDAD, LEY 97 o SIN NOMINA/NO VIGENTE |
| BIMESTRE | La etapa contiene "20" (leads dejados a futuro, p. ej. "ENERO 2027") |

**Columnas de Efectividad:**

* **Negocios cerrables** = Ganados + Oferta + Oferta en Espera + Documentación + Capturado + Lead Perdido.
* **Total asignación** = Negocios cerrables + Sin Capacidad + Bimestre.
* **% Efect. NETA** = Ganados / Negocios cerrables.
* **% Efect. BRUTA** = Ganados / Total asignación.

**Leads con más de un asesor en sus etiquetas:** se asignan al asesor que corresponde a su buzón (usuario responsable en Kommo). Si el buzón no corresponde a ninguno, se asignan al de la etiqueta puesta más recientemente. Si hay empate, cuentan para todos. La asignación se guarda en `leads_anomalos.json` y la usan los reportes.

**Ventas del día:** leads que llegaron por primera vez a una etapa de ganado el día FECHA_HASTA.

## Archivos generados

**Archivos del proceso** (en las rutas de `configuration.json`):

| Archivo | Contenido |
|---|---|
| `<CARPETA_BASE_ACUMULADOS>/<nombre>.xlsx` | Archivos base descargados de SharePoint. |
| `ACUMULADOS.xlsx` y `ACUMULADOS.info.json` | Hojas ACUMULADO CONTACTACION y ACUMULADO IA; periodo con que se generaron. |
| `historial_etapas_kommo.xlsx` y su `.info.json` | Hoja HISTORIAL: LEAD, EMBUDO, CREACION, ESTATUS DE NEGOCIO, FECHA DICTAMEN, MONTO OTORGADO, MOTIVO LEAD PERDIDO, ETIQUETAS, ETAPA ANTERIOR, ETAPA NUEVA, FECHA EVENTO, DIAS ETAPA ANTERIOR, BUZON. Hoja PIVOTE: un renglón por lead, con la primera fecha en cada "EMBUDO - ETAPA", EMBUDO ACTUAL y CAMBIOS DE EMBUDO. |
| `leads_anomalos.txt` / `.json` | Leads con varios asesores (fechas de etiquetas y asignación), con FECHA DICTAMEN pero sin ESTATUS, incompletos por motivo, y leads de ACUMULADOS no encontrados en Kommo. |

**Paquetes de reportes** (en la carpeta de reportes elegida), uno por prefijo:

```
<carpeta de reportes>/<FECHA_HASTA como AAAA-MM-DD>/
    GENERALES/
        <prefijo>reporte_general_dictaminados_<AAAA-MM-DD>.xlsx
    INDIVIDUALES/
        <prefijo>/
            <prefijo>reporte_individual_dictaminados_<asesor>.xlsx
```

Ejemplo con dos paquetes del mismo corte:

```
2026-09-30/
    GENERALES/
        solo_restringido_reporte_general_dictaminados_2026-09-30.xlsx
        sin_restringido_reporte_general_dictaminados_2026-09-30.xlsx
    INDIVIDUALES/
        solo_restringido_/   (un reporte por asesor)
        sin_restringido_/    (un reporte por asesor)
```

* **Reporte individual:** Resumen, Estado actual (con ruta del lead, buzón, origen y fecha de otorgado) e Historial por lead.
* **Reporte general:** Efectividad, Monto otorgado, Ventas del día y Detalle completo. Cada reporte general solo cuenta los reportes individuales de su paquete.

## Compilar el ejecutable

En Windows (PyInstaller solo genera programas para el sistema donde se ejecuta), con el entorno virtual activado y `json/secret.json` completo:

```bash
pip install pyinstaller
python compilar.py
```

Genera `dist/ReportesKommo/` (con `ReportesKommo.exe` y `_internal/`) y `dist/ReportesKommo.zip` para compartir. El `.exe` no funciona fuera de su carpeta.

## Seguridad

* `json/secret.json` y `preferencias.json` no deben subirse al repositorio (están en `.gitignore`).
* El ejecutable lleva dentro `secret.json`: el token de Kommo y los enlaces de SharePoint, cuyos archivos contienen datos personales de clientes. Compártelo solo con personas de confianza. Usa enlaces y un token que puedas revocar.
* Si un enlace o el token se llega a publicar, revócalo y genera uno nuevo.

## Estructura del proyecto

| Archivo | Función |
|---|---|
| `app.py` | Ventana del programa. |
| `descargar_base_acumulados.py` | Descarga los archivos base de SharePoint. |
| `generar_acumulados.py` | Genera ACUMULADOS a partir de los archivos base. |
| `lectura.py` | Lectura de hojas de Excel para el generador. |
| `acumulados.py` | Lectura de ACUMULADOS y de su periodo. |
| `extract_data_from_kommo.py` | Descarga de Kommo y armado del historial y de las anomalías. |
| `individual_reports.py` | Reportes individuales y clasificación de etapas. |
| `general_report.py` | Reporte general. |
| `functions.py` | Utilidades comunes (archivos, rutas, fechas, preferencias). |
| `setup_invironment.py` | Preparación del entorno. |
| `compilar.py` | Compilación del ejecutable. |
