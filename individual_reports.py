"""
Reporte de estado actual e historial de leads de Kommo.

Uso desde otro script:

    from individual_reports import generar_reporte_estado_leads

    generar_reporte_estado_leads(resultado_df, asesor="NOMBRE DEL ASESOR")

`resultado_df` es el resultado de la consulta: el historial de etapas de los
leads del asesor (filas de historial_etapas_kommo.xlsx). También acepta la
ruta a un .xlsx.

El reporte incluye tres hojas: Historial por lead, Resumen y Estado actual.
"""

import unicodedata
import json
import re
from datetime import datetime
from pathlib import Path

import duckdb as db
import pandas as pd

import functions as fn
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

CARPETA_REPORTES = "tablas/reportes/individuales"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Columnas mínimas que debe traer el resultado de la consulta.
COLUMNAS_REQUERIDAS = ["LEAD_ID", "ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA"]
# Columnas opcionales: si no vienen, se crean vacías.
COLUMNAS_OPCIONALES = ["DIAS_EN_ETAPA_ANTERIOR", "ETIQUETAS"]

# Categorías según la etapa actual (las mismas del reporte general), en el
# orden en que aparecen. Se comparan sin mayúsculas ni acentos con la lista
# de nombres de etapa equivalentes (de VENTAS y de CIERRES).
CATEGORIAS = [
    ("GANADOS", ["otorgado", "leads ganados"]),
    ("OFERTA", ["oferta"]),
    ("OFERTA EN ESPERA", ["oferta en espera"]),
    ("DOCUMENTACION", ["documentacion"]),
    ("CAPTURADO", ["capturado"]),
    ("LEAD PERDIDO", ["lead perdido", "no interesado"]),
    ("SIN CAPACIDAD", ["sin capacidad", "sin capacidad (cartera)"]),
]
CATEGORIA_GANADOS = "GANADOS"
CATEGORIA_FUTURO = "BIMESTRE"          # leads dejados a futuro: la etapa contiene "20"
CATEGORIA_OTROS = "Otros"              # ya no se usa: las etapas no clasificadas van a SIN CAPACIDAD
CATEGORIA_SIN_CAPACIDAD = "SIN CAPACIDAD"
CATEGORIA_PERDIDO = "LEAD PERDIDO"
# Como en el reporte de Cloud: un lead perdido con alguno de estos motivos
# (MOTIVO LEAD PERDIDO, aunque venga junto con otros) cuenta como SIN CAPACIDAD.
MOTIVOS_SIN_CAPACIDAD = ["sin capacidad", "ley 97", "sin nomina/no vigente"]
ORDEN_CATEGORIAS = [c for c, _ in CATEGORIAS] + [CATEGORIA_FUTURO]

_MAPA_ETAPAS = {alias: cat for cat, aliases in CATEGORIAS for alias in aliases}

FORMATO_FECHA = "dd/mm/yyyy hh:mm"
FORMATO_SOLO_FECHA = "dd/mm/yyyy"
# Nombre base de los archivos: <prefijo><NOMBRE_ARCHIVO_INDIVIDUAL>_<asesor>.xlsx.
# general_report.py y app.py lo importan de aquí; para renombrar, cambiarlo solo aquí.
NOMBRE_ARCHIVO_INDIVIDUAL = "reporte_individual_dictaminados"
FUENTE = "Arial"


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def normalizar(texto):
    """Minúsculas, sin acentos y con espacios simples."""
    if texto is None or pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def clasificar_etapa(etapa):
    """Devuelve la categoría de estado actual para una etapa de Kommo."""
    etapa_norm = normalizar(etapa)
    if etapa_norm in _MAPA_ETAPAS:
        return _MAPA_ETAPAS[etapa_norm]
    if es_dejado_a_futuro(etapa):
        return CATEGORIA_FUTURO
    return CATEGORIA_SIN_CAPACIDAD        # como Cloud: p. ej. "Contactado"


def etapa_no_clasificada(etapa):
    """True si la etapa no corresponde a ninguna categoría (cuenta como SIN CAPACIDAD)."""
    return normalizar(etapa) not in _MAPA_ETAPAS and not es_dejado_a_futuro(etapa)


def es_dejado_a_futuro(etapa):
    """DEJADO A FUTURO: la etapa contiene el texto "20" (p. ej. "ENERO 2027")."""
    return etapa is not None and not pd.isna(etapa) and "20" in str(etapa)


def es_ganado(etapa):
    """GANADO: la etapa es OTORGADO (de VENTAS o de CIERRES) o "Leads ganados"."""
    return clasificar_etapa(etapa) == CATEGORIA_GANADOS


COL_MOTIVO = "MOTIVO LEAD PERDIDO"
COL_MONTO = "MONTO OTORGADO"
FORMATO_DINERO = '"$"#,##0.00'


def categoria_lead(etapa, motivo=None):
    """Categoría de un lead según su etapa actual, con la regla de Cloud: un lead
    perdido cuyo MOTIVO LEAD PERDIDO incluye SIN CAPACIDAD, LEY 97 o SIN
    NOMINA/NO VIGENTE cuenta como SIN CAPACIDAD."""
    categoria = clasificar_etapa(etapa)
    if categoria == CATEGORIA_PERDIDO and motivo is not None and not pd.isna(motivo):
        texto = normalizar(motivo)
        if any(m in texto for m in MOTIVOS_SIN_CAPACIDAD):
            return CATEGORIA_SIN_CAPACIDAD
    return categoria


# Títulos de la hoja HISTORIAL (extract_data_from_kommo.py) -> nombres que
# usan los reportes y la consulta de secret.json. Los demás no cambian.
COLUMNAS_HISTORIAL_INTERNAS = {
    # V11
    "LEAD": "LEAD_ID",
    "CREACION": "creacion_de_lead",
    "ETAPA NUEVA": "ETAPA_NUEVA",
    "DIAS ETAPA ANTERIOR": "DIAS_EN_ETAPA_ANTERIOR",
    # V10
    "LEAD ID": "LEAD_ID",
    "FECHA CREACION DE LEAD": "creacion_de_lead",
    "ETAPA ANTERIOR": "ETAPA_ANTERIOR",
    "FECHA EVENTO": "FECHA",
    "DIAS EN ETAPA ANTERIOR": "DIAS_EN_ETAPA_ANTERIOR",
}


def normalizar_columnas_historial(df):
    """Traduce los títulos del HISTORIAL a los nombres internos.

    Acepta también historiales con los nombres anteriores (no los toca).
    """
    return df.rename(columns={k: v for k, v in COLUMNAS_HISTORIAL_INTERNAS.items()
                              if k in df.columns})


def formato_tiempo(dias):
    """Convierte días decimales a texto legible, p. ej. 12.4 -> '12 d 10 h'."""
    if dias is None or pd.isna(dias):
        return ""
    total_horas = int(round(dias * 24))
    d, h = divmod(total_horas, 24)
    return f"{d} d {h} h"


def etapa_con_embudo(etapa, embudo=None):
    """'Oferta' + 'CIERRES' -> 'CIERRES: Oferta'. Sin embudo, solo la etapa."""
    if etapa is None or pd.isna(etapa):
        return None
    if embudo is None or pd.isna(embudo) or not str(embudo).strip():
        return str(etapa)
    return f"{embudo}: {etapa}"


def construir_ruta(grupo):
    """Ruta de un lead con embudo y etapa:
    '(sin etapa previa) → VENTAS: Nuevo (01/09/2026 09:20) → CIERRES: Oferta (...)'."""
    embudos = grupo["EMBUDO"] if "EMBUDO" in grupo else [None] * len(grupo)
    primera_anterior = grupo["ETAPA_ANTERIOR"].iloc[0]
    emb_anterior = grupo["EMBUDO_ANTERIOR"].iloc[0] if "EMBUDO_ANTERIOR" in grupo else None
    partes = [etapa_con_embudo(primera_anterior, emb_anterior) if pd.notna(primera_anterior)
              else "(sin etapa previa)"]
    for etapa, fecha, embudo in zip(grupo["ETAPA_NUEVA"], grupo["FECHA"], embudos):
        partes.append(f"{etapa_con_embudo(etapa, embudo)} ({fecha:%d/%m/%Y %H:%M})")
    return " → ".join(partes)


# ---------------------------------------------------------------------------
# Procesamiento
# ---------------------------------------------------------------------------
def filtrar_por_etiqueta(df, etiqueta):
    """
    Devuelve el historial completo de los leads que tienen la etiqueta en
    alguno de sus registros. La comparación no distingue mayúsculas ni acentos.
    """
    con = db.connect()
    con.register("historial", df)
    query = """
        SELECT * FROM historial
        WHERE LEAD_ID IN (
            SELECT LEAD_ID FROM historial
            WHERE strip_accents(ETIQUETAS) ILIKE strip_accents($1)
        )
    """
    resultado = con.execute(query, [f"%{etiqueta}%"]).df()
    con.close()
    return resultado


def preparar_historial(df, fecha_corte):
    """Ordena el historial y agrega paso, último movimiento y días en cada etapa."""
    con = db.connect()
    con.register("historial", df)
    hist = con.execute("""
        SELECT *,
            ROW_NUMBER() OVER w                 AS PASO,
            COUNT(*) OVER (PARTITION BY LEAD_ID) AS TOTAL_MOVIMIENTOS,
            LEAD(FECHA) OVER w                  AS FECHA_SIGUIENTE,
            LAG(EMBUDO) OVER w                  AS EMBUDO_ANTERIOR
        FROM historial
        WINDOW w AS (PARTITION BY LEAD_ID ORDER BY FECHA)
        ORDER BY LEAD_ID, FECHA
    """).df()
    con.close()

    hist["FECHA"] = pd.to_datetime(hist["FECHA"])
    hist["FECHA_SIGUIENTE"] = pd.to_datetime(hist["FECHA_SIGUIENTE"])
    hist["ES_ULTIMO_MOVIMIENTO"] = hist["PASO"] == hist["TOTAL_MOVIMIENTOS"]

    # Días en la etapa nueva: hasta el siguiente movimiento, o hasta la fecha
    # de corte si es la etapa actual.
    fin_etapa = hist["FECHA_SIGUIENTE"].fillna(pd.Timestamp(fecha_corte))
    hist["DIAS_EN_ETAPA_NUEVA"] = ((fin_etapa - hist["FECHA"]).dt.total_seconds() / 86400).round(2)
    return hist


def generar_reporte(df, fecha_corte):
    """Devuelve un dict con los DataFrames de cada hoja del reporte."""
    df = df.copy()
    df["FECHA"] = pd.to_datetime(df["FECHA"])
    if "EMBUDO" not in df.columns:      # historial sin embudo: las rutas van sin él
        df["EMBUDO"] = None
    hist = preparar_historial(df, fecha_corte)
    for col in COLUMNAS_OPCIONALES:
        if col not in hist.columns:
            hist[col] = None

    # --- Estado actual: último movimiento de cada lead -----------------------
    ultimos = hist[hist["ES_ULTIMO_MOVIMIENTO"]].copy()
    motivos = ultimos[COL_MOTIVO] if COL_MOTIVO in ultimos.columns else [None] * len(ultimos)
    ultimos["ESTADO_ACTUAL"] = [categoria_lead(e, m) for e, m in zip(ultimos["ETAPA_NUEVA"], motivos)]
    ultimos["DIAS_DESDE_ULTIMO_MOVIMIENTO"] = (
        (pd.Timestamp(fecha_corte) - ultimos["FECHA"]).dt.total_seconds() / 86400
    ).round(2)
    ultimos["TIEMPO_TRANSCURRIDO"] = ultimos["DIAS_DESDE_ULTIMO_MOVIMIENTO"].map(formato_tiempo)

    # Las etiquetas pueden venir solo en algunos registros: se toma la última no vacía.
    etiquetas = hist.groupby("LEAD_ID")["ETIQUETAS"].last()
    rutas = hist.groupby("LEAD_ID")[["ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA", "EMBUDO",
                                     "EMBUDO_ANTERIOR"]].apply(construir_ruta)
    ultimos["ETIQUETAS"] = ultimos["LEAD_ID"].map(etiquetas)
    ultimos["RUTA"] = ultimos["LEAD_ID"].map(rutas)
    ultimos["GANADO"] = ultimos["ETAPA_NUEVA"].map(es_ganado).astype(bool)
    # Primera vez que el lead llegó a una etapa de ganado (OTORGADO / Leads ganados)
    ganados = hist[hist["ETAPA_NUEVA"].map(es_ganado)]
    primera = ganados.groupby("LEAD_ID")["FECHA"].min().to_dict()
    ultimos["FECHA_OTORGADO"] = pd.to_datetime(ultimos["LEAD_ID"].map(primera), errors="coerce")
    for col in ("BUZON", "ORIGEN"):            # dato del lead: el último no vacío
        if col in hist.columns:
            ultimos[col] = ultimos["LEAD_ID"].map(hist.groupby("LEAD_ID")[col].last())
    ultimos["DEJADO A FUTURO"] = ultimos["ETAPA_NUEVA"].map(es_dejado_a_futuro).astype(bool)

    ultimos["_orden"] = ultimos["ESTADO_ACTUAL"].map(ORDEN_CATEGORIAS.index)
    ultimos = ultimos.sort_values(["_orden", "DIAS_DESDE_ULTIMO_MOVIMIENTO"], ascending=[True, False])

    columnas_estado = [
        "LEAD_ID", "creacion_de_lead","ESTADO_ACTUAL", "ETAPA_NUEVA", "FECHA",
        "DIAS_DESDE_ULTIMO_MOVIMIENTO", "TIEMPO_TRANSCURRIDO", "ETAPA_ANTERIOR",
        "ESTATUS DE NEGOCIO", "FECHA DICTAMEN", COL_MOTIVO, "TOTAL_MOVIMIENTOS", "RUTA",
        "ETIQUETAS", "EMBUDO", "GANADO", "DEJADO A FUTURO", COL_MONTO,
        "BUZON", "ORIGEN", "FECHA_OTORGADO",
    ]
    estado = (
        ultimos[[c for c in columnas_estado if c in ultimos.columns]]
        .rename(columns={"ETAPA_NUEVA": "ETAPA_ACTUAL", "FECHA": "FECHA_ULTIMO_MOVIMIENTO",
                         "EMBUDO": "EMBUDO_ACTUAL"})
        .reset_index(drop=True)
    )

    # --- Resumen por categoría ----------------------------------------------
    total = len(estado)
    filas = []
    for cat in ORDEN_CATEGORIAS:
        sub = estado[estado["ESTADO_ACTUAL"] == cat]
        filas.append({
            "ESTADO_ACTUAL": cat,
            "LEADS": len(sub),
            "PORCENTAJE": len(sub) / total if total else 0,
            "PROMEDIO_DIAS_DESDE_ULTIMO_MOVIMIENTO": round(sub["DIAS_DESDE_ULTIMO_MOVIMIENTO"].mean(), 2) if len(sub) else None,
        })
    filas.append({
        "ESTADO_ACTUAL": "TOTAL",
        "LEADS": total,
        "PORCENTAJE": 1 if total else 0,
        "PROMEDIO_DIAS_DESDE_ULTIMO_MOVIMIENTO": round(estado["DIAS_DESDE_ULTIMO_MOVIMIENTO"].mean(), 2) if total else None,
    })
    resumen = pd.DataFrame(filas)

    # --- Historial ------------------------------------------------------------
    hist["ESTADO_ACTUAL_LEAD"] = hist["LEAD_ID"].map(estado.set_index("LEAD_ID")["ESTADO_ACTUAL"])
    hist["ES_ULTIMO_MOVIMIENTO"] = hist["ES_ULTIMO_MOVIMIENTO"].map({True: "Sí", False: ""})
    historial = hist[[
        "LEAD_ID", "PASO", "ETAPA_ANTERIOR", "ETAPA_NUEVA", "FECHA",
        "DIAS_EN_ETAPA_ANTERIOR", "DIAS_EN_ETAPA_NUEVA",
        "ESTADO_ACTUAL_LEAD", "ES_ULTIMO_MOVIMIENTO",
        "EMBUDO", "EMBUDO_ANTERIOR",          # para mostrar "EMBUDO: etapa" en los bloques
    ]].reset_index(drop=True)

    return {"Resumen": resumen, "Estado actual": estado, "Historial": historial}


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
_RELLENO_ENCABEZADO = PatternFill("solid", start_color="1F3864")
_ANCHOS = {"RUTA": 90, "ETIQUETAS": 60}


def _dar_formato(ws, fila_encabezado, formatos=None):
    """Fuente, encabezado, anchos, filtros y paneles fijos para una hoja."""
    formatos = formatos or {}
    encabezados = [c.value for c in ws[fila_encabezado]]

    for fila in ws.iter_rows():
        for celda in fila:
            celda.font = Font(name=FUENTE, bold=celda.font.bold)

    for celda in ws[fila_encabezado]:
        celda.font = Font(name=FUENTE, bold=True, color="FFFFFF")
        celda.fill = _RELLENO_ENCABEZADO
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, nombre in enumerate(encabezados, start=1):
        letra = ws.cell(row=fila_encabezado, column=idx).column_letter
        if nombre in formatos:
            for (celda,) in ws.iter_rows(min_row=fila_encabezado + 1, min_col=idx, max_col=idx):
                celda.number_format = formatos[nombre]
        largo = max(
            (len(str(c.value)) for (c,) in ws.iter_rows(min_row=fila_encabezado, min_col=idx, max_col=idx) if c.value is not None),
            default=10,
        )
        ws.column_dimensions[letra].width = _ANCHOS.get(nombre, min(max(largo + 2, 12), 40))

    ws.row_dimensions[fila_encabezado].height = 30
    ws.freeze_panes = ws.cell(row=fila_encabezado + 1, column=2)
    ws.auto_filter.ref = f"A{fila_encabezado}:{ws.cell(row=ws.max_row, column=len(encabezados)).coordinate}"


_RELLENO_LEAD = PatternFill("solid", start_color="D9E1F2")
_RELLENO_SUBENCABEZADO = PatternFill("solid", start_color="8EA9DB")
HOJA_POR_LEAD = "Historial por lead"
# Columnas de la hoja "Estado actual": (nombre interno, título en el Excel),
# en orden. El DataFrame interno conserva más datos porque los usan el
# Resumen y el Historial por lead. general_report.py lee estos títulos.
COLUMNAS_HOJA_ESTADO = [
    ("LEAD_ID", "LEAD"),
    ("creacion_de_lead", "CREACION"),
    ("EMBUDO_ACTUAL", "EMBUDO"),
    ("ETAPA_ACTUAL", "ETAPA ACTUAL"),
    ("ETIQUETAS", "ETIQUETAS"),
    ("ESTATUS DE NEGOCIO", "ESTATUS DE NEGOCIO"),
    ("FECHA DICTAMEN", "FECHA DICTAMEN"),
    ("GANADO", "GANADO"),
    ("DEJADO A FUTURO", "DEJADO A FUTURO"),
    (COL_MONTO, COL_MONTO),
    (COL_MOTIVO, COL_MOTIVO),
    ("FECHA_ULTIMO_MOVIMIENTO", "FECHA ULTIMO MOVIMIENTO"),
    ("TIEMPO_TRANSCURRIDO", "TIEMPO TRANSCURRIDO"),
    ("TOTAL_MOVIMIENTOS", "TOTAL MOVIMIENTOS"),
    ("RUTA", "RUTA LEAD"),
    # agregadas al final para no mover las anteriores
    ("BUZON", "BUZON"),
    ("ORIGEN", "ORIGEN"),
    ("FECHA_OTORGADO", "FECHA OTORGADO"),
]
COLUMNAS_BLOQUE = [
    ("PASO", "PASO", 8, None),
    ("ETAPA_ANTERIOR", "ETAPA ANTERIOR", 22, None),
    ("ETAPA_NUEVA", "ETAPA NUEVA", 22, None),
    ("FECHA", "FECHA DEL CAMBIO", 20, FORMATO_FECHA),
    ("DIAS_EN_ETAPA_ANTERIOR", "DÍAS EN ETAPA ANTERIOR", 16, "0.00"),
    ("DIAS_EN_ETAPA_NUEVA", "DÍAS EN ETAPA NUEVA", 16, "0.00"),
]


def _escribir_historial_por_lead(libro, estado, historial):
    """
    Una hoja con un bloque por lead: encabezado con el lead, su etapa actual,
    la fecha de su última actualización de etapa y el tiempo transcurrido desde
    entonces; debajo, la tabla con todos sus cambios de etapa en orden.
    Devuelve {LEAD_ID: fila donde empieza su bloque} para crear vínculos.
    """
    ws = libro.create_sheet(HOJA_POR_LEAD, 0)
    n_cols = len(COLUMNAS_BLOQUE)
    negrita = Font(name=FUENTE, bold=True)
    normal = Font(name=FUENTE)
    inicio_bloques = {}
    fila = 1

    for _, lead in estado.iterrows():
        lead_id = lead["LEAD_ID"]
        inicio_bloques[lead_id] = fila

        # Fila 1 del bloque: identificación del lead
        actual = etapa_con_embudo(lead["ETAPA_ACTUAL"], lead.get("EMBUDO_ACTUAL"))
        titulo = f"LEAD {lead_id}  |  ETAPA ACTUAL: {str(actual).upper()}"
        ws.cell(row=fila, column=1, value=titulo).font = Font(name=FUENTE, bold=True, color="FFFFFF", size=12)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=n_cols)
        for col in range(1, n_cols + 1):
            ws.cell(row=fila, column=col).fill = _RELLENO_ENCABEZADO
        fila += 1

        # Fila 2: última actualización de etapa y tiempo transcurrido
        ws.cell(row=fila, column=1, value="Última actualización").font = negrita
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
        c = ws.cell(row=fila, column=3, value=lead["FECHA_ULTIMO_MOVIMIENTO"].to_pydatetime())
        c.font = normal
        c.number_format = FORMATO_FECHA
        c.alignment = Alignment(horizontal="left")
        ws.cell(row=fila, column=4, value="Tiempo transcurrido").font = negrita
        ws.cell(
            row=fila, column=5,
            value=f"{lead['TIEMPO_TRANSCURRIDO']}  ({lead['DIAS_DESDE_ULTIMO_MOVIMIENTO']:.2f} días)",
        ).font = negrita
        ws.merge_cells(start_row=fila, start_column=5, end_row=fila, end_column=n_cols)
        for col in range(1, n_cols + 1):
            ws.cell(row=fila, column=col).fill = _RELLENO_LEAD
        fila += 1

        # Encabezado de la tabla de cambios
        for col, (_, titulo_col, _, _) in enumerate(COLUMNAS_BLOQUE, start=1):
            c = ws.cell(row=fila, column=col, value=titulo_col)
            c.font = Font(name=FUENTE, bold=True, color="FFFFFF")
            c.fill = _RELLENO_SUBENCABEZADO
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        fila += 1

        # Cambios de etapa en orden cronológico; el más reciente en negritas
        movimientos = historial[historial["LEAD_ID"] == lead_id].sort_values("PASO")
        for _, mov in movimientos.iterrows():
            es_ultimo = mov["ES_ULTIMO_MOVIMIENTO"] == "Sí"
            for col, (campo, _, _, fmt) in enumerate(COLUMNAS_BLOQUE, start=1):
                valor = mov[campo]
                if campo == "ETAPA_NUEVA":          # con su embudo: "CIERRES: Oferta"
                    valor = etapa_con_embudo(valor, mov.get("EMBUDO"))
                elif campo == "ETAPA_ANTERIOR":
                    valor = etapa_con_embudo(valor, mov.get("EMBUDO_ANTERIOR"))
                if valor is None or pd.isna(valor):
                    valor = None
                elif campo == "FECHA":
                    valor = valor.to_pydatetime()
                c = ws.cell(row=fila, column=col, value=valor)
                c.font = negrita if es_ultimo else normal
                if fmt:
                    c.number_format = fmt
                if campo == "PASO":
                    c.alignment = Alignment(horizontal="center")
            fila += 1

        fila += 1  # fila en blanco entre leads

    for col, (_, _, ancho, _) in enumerate(COLUMNAS_BLOQUE, start=1):
        ws.column_dimensions[get_column_letter(col)].width = ancho
    ws.sheet_view.showGridLines = False
    return inicio_bloques


def _formato_fecha_sin_hora(ws, encabezado):
    """Formato dd/mm/yyyy para una columna de fechas; si una fecha trae hora, se muestra."""
    for idx, celda in enumerate(ws[1], start=1):
        if celda.value != encabezado:
            continue
        for (c,) in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            if hasattr(c.value, "hour"):
                con_hora = (c.value.hour, c.value.minute, c.value.second) != (0, 0, 0)
                c.number_format = FORMATO_FECHA if con_hora else FORMATO_SOLO_FECHA


def escribir_excel(hojas, ruta, fecha_corte, asesor=""):
    """Escribe el reporte con fechas reales de Excel y formato legible."""
    with pd.ExcelWriter(ruta, engine="openpyxl", datetime_format=FORMATO_FECHA) as writer:
        fila_tabla_resumen = 5
        hojas["Resumen"].to_excel(writer, sheet_name="Resumen", index=False, startrow=fila_tabla_resumen - 1)
        estado = hojas["Estado actual"].copy()
        for interno, _ in COLUMNAS_HOJA_ESTADO:     # columnas que no vengan: vacías
            if interno not in estado.columns:
                estado[interno] = None
        (estado[[interno for interno, _ in COLUMNAS_HOJA_ESTADO]]
         .rename(columns=dict(COLUMNAS_HOJA_ESTADO))
         .to_excel(writer, sheet_name="Estado actual", index=False))

        libro = writer.book

        # --- Historial por lead (primera hoja)
        inicio_bloques = _escribir_historial_por_lead(libro, hojas["Estado actual"], hojas["Historial"])
        libro.active = 0

        # --- Resumen
        ws = libro["Resumen"]
        ws["A1"], ws["B1"] = "Fecha de corte", fecha_corte
        ws["B1"].number_format = FORMATO_FECHA
        ws["A2"], ws["B2"] = "Asesor", asesor
        _dar_formato(ws, fila_tabla_resumen, {"PORCENTAJE": "0.0%", "PROMEDIO_DIAS_DESDE_ULTIMO_MOVIMIENTO": "0.00"})
        ws.auto_filter.ref = None
        ws.freeze_panes = None
        for c in ("A1", "A2"):
            ws[c].font = Font(name=FUENTE, bold=True)
        for celda in ws[ws.max_row]:  # fila TOTAL
            celda.font = Font(name=FUENTE, bold=True)
        nota = ws.max_row + 2
        ws.cell(row=nota, column=1, value="Notas:").font = Font(name=FUENTE, bold=True)
        notas = [
            "El tiempo transcurrido se calcula desde la última actualización de etapa hasta la fecha de corte (momento en que se generó el reporte).",
            "Las categorías dependen de la etapa actual del lead, igual que en el reporte general.",
            "'BIMESTRE' agrupa los leads dejados a futuro (la etapa contiene \"20\", p. ej. \"ENERO 2027\").",
            "'SIN CAPACIDAD' incluye las etapas no clasificadas (p. ej. Contactado) y los leads perdidos por SIN CAPACIDAD, LEY 97 o SIN NOMINA/NO VIGENTE.",
            "En 'Historial por lead', los días en la etapa nueva del último cambio se cuentan hasta la fecha de corte.",
        ]
        for i, texto in enumerate(notas, start=1):
            ws.cell(row=nota + i, column=1, value=texto).font = Font(name=FUENTE)
        ws.column_dimensions["A"].width = 22

        # --- Estado actual, con vínculo de cada LEAD_ID a su bloque de historial
        ws = libro["Estado actual"]
        _dar_formato(ws, 1, {
            "CREACION": FORMATO_FECHA,
            "FECHA ULTIMO MOVIMIENTO": FORMATO_FECHA,
            "FECHA OTORGADO": FORMATO_FECHA,
            COL_MONTO: FORMATO_DINERO,
        })
        _formato_fecha_sin_hora(ws, "FECHA DICTAMEN")
        for (celda,) in ws.iter_rows(min_row=2, max_col=1):
            if celda.value in inicio_bloques:
                celda.hyperlink = Hyperlink(
                    ref=celda.coordinate, location=f"'{HOJA_POR_LEAD}'!A{inicio_bloques[celda.value]}"
                )
                celda.font = Font(name=FUENTE, color="0563C1", underline="single")


# ---------------------------------------------------------------------------
# Función pública
# ---------------------------------------------------------------------------
ARCHIVO_ANOMALIAS_JSON = "leads_anomalos.json"


def cargar_asignaciones(carpeta):
    """{LEAD_ID: asesor asignado} de leads_anomalos.json (en la carpeta del
    historial). Solo los leads con un asesor asignado; los empates no aparecen
    (cuentan para todos sus asesores). Sin archivo: {}."""
    ruta = Path(carpeta) / ARCHIVO_ANOMALIAS_JSON
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return {int(d["lead"]): d["asignado"] for d in datos.get("varias_asesores") or []
            if d.get("asignado")}


def aplicar_asignacion(historial, asesor, asignaciones):
    """Quita del historial de un asesor los leads asignados a OTRO asesor."""
    if not asignaciones or historial is None or historial.empty:
        return historial
    ajenos = {lid for lid, a in asignaciones.items() if normalizar(a) != normalizar(asesor)}
    return historial[~historial["LEAD_ID"].isin(ajenos)]


def _carpeta_por_fecha(carpeta_base, fecha):
    """tablas + 23/09/2026 -> tablas/septiembre/23"""
    return Path(carpeta_base) / MESES[fecha.month - 1] / f"{fecha.day:02d}"

def _nombre_archivo(asesor, prefijo=""):
    """'Ana Prueba' -> 'reporte_individual_dictaminados_ana_prueba.xlsx'.

    El prefijo, si se indica, va al inicio: 'ENERO_reporte_individual_dictaminados_...'.
    """
    slug = re.sub(r"[^a-z0-9]+", "_", normalizar(asesor)).strip("_")
    nombre = f"{NOMBRE_ARCHIVO_INDIVIDUAL}_{slug}" if slug else NOMBRE_ARCHIVO_INDIVIDUAL
    return f"{prefijo}{nombre}.xlsx"


def generar_reporte_estado_leads(historial, asesor="", ruta_salida=None,
                                 fecha_corte=None, carpeta_salida=None,
                                 verbose=True, fecha_archivo=None, prefijo=""):
    """
    Genera el reporte de estado actual e historial de leads en Excel.

    Parámetros
    ----------
    historial : DataFrame o ruta (str/Path) a un .xlsx
        Resultado de la consulta: historial de etapas de los leads del asesor.
        Debe traer LEAD_ID, ETAPA_ANTERIOR, ETAPA_NUEVA y FECHA.
    asesor : str
        Nombre del asesor. Se muestra en el Resumen y se usa para nombrar el archivo.
    ruta_salida : str o Path, opcional
        Ruta completa del Excel. Si no se indica, se usa
        <carpeta_salida>/<mes>/<día>/reporte_individual_dictaminados_<asesor>.xlsx.
    fecha_corte : datetime, opcional
        Fecha contra la que se calcula el tiempo transcurrido. Por defecto, ahora.
        Útil para que todos los reportes de una misma corrida usen la misma fecha.
    carpeta_salida : str o Path, opcional
        Carpeta donde se guarda el reporte cuando no se indica ruta_salida.
        Por defecto, tablas/reportes/individuales dentro del proyecto; el
        reporte queda en <carpeta_salida>/<mes>/<día>/.
    verbose : bool
        Si es True, imprime un resumen en consola.

    Devuelve
    --------
    Path del archivo generado, o None si la consulta no trajo registros.
    """
    if isinstance(historial, (str, Path)):
        historial = pd.read_excel(historial)
    historial = normalizar_columnas_historial(historial)

    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in historial.columns]
    if faltantes:
        raise ValueError(f"A la consulta le faltan columnas requeridas: {faltantes}")

    if historial.empty:
        if verbose:
            print(f"[{asesor or 'sin asesor'}] La consulta no trajo registros; no se generó reporte.")
        return None

    fecha_corte = fecha_corte or datetime.now().replace(microsecond=0)
    if carpeta_salida is None:
        carpeta_salida = fn.ruta_proyecto(CARPETA_REPORTES)
    if ruta_salida:
        ruta_salida = Path(ruta_salida)
    else:
        fecha_archivo = fecha_archivo or fn.fecha_reportes()
        ruta_salida = _carpeta_por_fecha(carpeta_salida, fecha_archivo) / _nombre_archivo(asesor, prefijo=prefijo)
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    hojas = generar_reporte(historial, fecha_corte)
    escribir_excel(hojas, ruta_salida, fecha_corte, asesor)

    if verbose:
        print(f"[{asesor or 'sin asesor'}] Reporte generado: {fn.ruta_para_mostrar(ruta_salida)} "
              f"({len(hojas['Estado actual'])} leads, {len(hojas['Historial'])} movimientos)")
    return ruta_salida


# ---------------------------------------------------------------------------
# Ejecución directa: ejemplo con un asesor
# ---------------------------------------------------------------------------
def main():
    labels = fn.import_json("json/secret.json")
    asesor = labels["asesores"][0]  # primer asesor de la lista

    ruta = fn.ruta_historial()                      # configuration.json -> CARPETA_HISTORIAL
    df = normalizar_columnas_historial(fn.import_xlsx(ruta))
    historial = aplicar_asignacion(filtrar_por_etiqueta(df, asesor), asesor,
                                   cargar_asignaciones(fn.carpeta_anomalos()))
    generar_reporte_estado_leads(historial, asesor=asesor)


if __name__ == "__main__":
    main()