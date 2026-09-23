import functions as fn
import pandas as pd
import duckdb as db

labels = fn.import_json("json/secret.json")

df = fn.import_xlsx("tablas/historial_etapas_kommo 2.xlsx")

query = f"""
                SELECT DISTINCT LEAD_ID FROM df 
                WHERE 
                    ETIQUETAS ILIKE '%{labels["people"]["sin_financiera_restringida"][0]}%'
                """

print(db.query(query).df())

#print (f"{labels['people']['sin_financiera_restringida'][0]}")