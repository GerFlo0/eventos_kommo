import functions as fn
import duckdb as db

labels = fn.import_json("json/secret.json")

df = fn.import_xlsx("tablas/historial_etapas_kommo.xlsx")

# 1. Definir la consulta usando un marcador de posición ($1)
query = """
    SELECT * FROM df 
    WHERE 
        LEAD_ID IN (
            SELECT LEAD_ID FROM df 
            WHERE 
                ETIQUETAS LIKE $1
        )
"""

# 2. Preparar el valor agregando los comodines % antes y después
filtro_etiqueta = f'%{labels["people"]["sin_financiera_restringida"][0]}%' #Miriam Gomez Cierres

# 3. Pasar el parámetro en el método .execute() o directamente en .query()
resultado_df = db.execute(query, [filtro_etiqueta]).df()

#print (f"{labels['people']['sin_financiera_restringida'][0]}")

#save the result to a new Excel file
resultado_df.to_excel("result.xlsx", index=False)

#print result
#print(resultado_df)