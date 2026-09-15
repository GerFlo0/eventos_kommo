# Decargador de eventos de kommo

Descarga los eventos de leads en un rango de fechas de kommo y genera excel donde se muestra cuanto tiempo estuvo un lead en cada etapa

## Requisitos previos

* Tener instalado Python (versión 3.8 o superior recomendada).
* Tener instalado pip (gestor de paquetes de Python).

## Instalación y Configuración

Sigue estos pasos para configurar el entorno de desarrollo en tu computadora:

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/GerFlo0/eventos_kommo.git
   ```

2. Crear el entorno virtual:
   * En Windows:
     ```bash
     python -m venv .venv
     ```
   * En macOS / Linux:
     ```bash
     python3 -m venv .venv
     ```

3. Activar el entorno virtual:
   * En Windows (CMD o PowerShell):
     ```bash
     .venv\Scripts\activate
     ```
   * En macOS / Linux:
     ```bash
     source .venv/bin/activate
     ```

4. Instalar las dependencias:
   Instala los paquetes necesarios listados en el archivo requirements.txt:
   ```bash
   pip install -r requirements.txt
   ```

## Uso

modifica los parametros en json configuration segun tus necesidades y ejecuta el programa

```bash
python main.py
```
