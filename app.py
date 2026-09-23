from datetime import datetime
import duckdb as db
import functions as fn
from individual_reports import generar_reporte_estado_leads

secret = fn.import_json("json/secret.json")
df = fn.import_xlsx("tablas/historial_etapas_kommo.xlsx")
fecha_corte = datetime.now().replace(microsecond=0)  # misma fecha para todos los reportes

for asesor in secret["people"]["sin_financiera_restringida"]:
    resultado_df = db.execute(secret["query"], [f'%{asesor}%']).df()  # tu consulta para este asesor
    generar_reporte_estado_leads(resultado_df, asesor=asesor, fecha_corte=fecha_corte)