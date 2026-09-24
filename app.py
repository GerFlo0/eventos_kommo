"""
Genera un reporte individual por cada asesor de la lista
secret.json -> people -> sin_financiera_restringida.

Requiere haber corrido antes extract_data_from_kommo.py (lee
tablas/historial_etapas_kommo.xlsx). La consulta de cada asesor está en
secret.json -> query y se ejecuta sobre la tabla "df" (el historial).
"""
from datetime import datetime

import duckdb as db

import functions as fn
from individual_reports import generar_reporte_estado_leads


def main():
    secret = fn.import_json("json/secret.json")
    df = fn.import_xlsx("tablas/historial_etapas_kommo.xlsx")
    fecha_corte = datetime.now().replace(microsecond=0)  # misma fecha para todos los reportes

    # La consulta lee la tabla "df". Se registra explícitamente en lugar de
    # depender de que DuckDB encuentre la variable de Python por su nombre.
    con = db.connect()
    con.register("df", df)
    try:
        for asesor in secret["people"]["sin_financiera_restringida"]:
            resultado_df = con.execute(secret["query"], [f'%{asesor}%']).df()  # consulta de este asesor
            generar_reporte_estado_leads(resultado_df, asesor=asesor, fecha_corte=fecha_corte)
    finally:
        con.close()


if __name__ == "__main__":
    main()
