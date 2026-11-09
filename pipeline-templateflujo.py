import os
import pandas as pd
import psycopg2
import requests
from dotenv import load_dotenv

# Fuerza a buscar el .env en tu carpeta real del Escritorio
RUTA_ENV = r"C:\Users\camila.moraga\Desktop\ESTRATEGIA\visual\.venv\.env"
load_dotenv(RUTA_ENV)

HOST = os.environ["REDSHIFT_HOST"]
DBNAME = os.environ["REDSHIFT_DBNAME"]
USER = os.environ["REDSHIFT_USER"]
PASSWORD = os.environ["REDSHIFT_PASSWORD"]
PORT = int(os.environ["REDSHIFT_PORT"])

# Carga de las dos URLs de Google Sheets
WEBHOOK_URL1 = os.environ["WEBHOOK_URL1"]

# 2. Consulta SQL
QUERY = """
WITH base_preparada AS (
    SELECT 
        date,
        in_reported,
        id_platform,
        CASE id_property
            WHEN 'MAP' THEN 'ASA'
            WHEN 'PLZ' THEN 'MAE'
            WHEN 'MQU' THEN 'MAQ'
            WHEN 'BAV' THEN 'BOM'
            ELSE id_property
        END AS id_property,
        LAST_VALUE(entrance_name) OVER(
            PARTITION BY id_property, entrance_key
            ORDER BY date ASC
            ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
        ) AS entrance_name
    FROM parauco_analysis_opscentroscomerciales.agg_entsal_footfall_flow_cl
    WHERE (id_platform = 'REM' OR id_platform = 'SOM')
      AND DATE_PART('year', date) IN (2025, 2026)
      AND date <= CURRENT_DATE
)
SELECT
    TO_CHAR(date, 'DD-MM-YYYY') AS date_formateada,
    DATE_PART('year', date) AS year,
    id_property,
    entrance_name,
    id_platform,
    SUM(in_reported) AS flujo,
    DATE_PART('week', date) AS semana,
    CASE LOWER(REPLACE(TO_CHAR(date, 'TMDay'), 'TM', ''))
        WHEN 'wednesday' THEN 'miercoles'
        WHEN 'monday'    THEN 'lunes'
        WHEN 'tuesday'   THEN 'martes'
        WHEN 'thursday'  THEN 'jueves'
        WHEN 'friday'    THEN 'viernes'
        WHEN 'saturday'  THEN 'sabado'
        WHEN 'sunday'    THEN 'domingo'
        ELSE LOWER(REPLACE(TO_CHAR(date, 'TMDay'), 'TM', ''))
    END AS dia_nombre,
    DATE_PART('month', date) AS mes,
    TO_CHAR(
        CASE 
            WHEN DATE_PART('year', date) = 2025 THEN DATEADD(day, 364, date)::DATE
            ELSE date
        END,
        'DD-MM-YYYY'
    ) AS fecha_homologada
FROM base_preparada
GROUP BY 1, 2, 3, 4, 5, 7, 8, 9, 10
ORDER BY date_formateada ASC, id_property ASC;
"""
def ejecutar_y_enviar():
    print("⏳ Conectando a Redshift y ejecutando consulta...")
    conn = psycopg2.connect(
        host=HOST, dbname=DBNAME, user=USER, password=PASSWORD, port=PORT
    )
    
    df = pd.read_sql(QUERY, conn)
    conn.close()
    print(f"✅ Consulta procesada con éxito: {len(df)} filas.")

    # 1. Matriz para Estimaciones_SOM (mantiene id_platform o todas las columnas necesarias)
    df_som = df[df['id_platform'] == 'SOM']
    matriz_som = [df_som.columns.tolist()] + df_som.fillna("").values.tolist()

    # 2. Matriz para BBDD: Eliminamos la columna 'id_platform' para mantener las 10 columnas exactas
    df_bbdd = df.drop(columns=['id_platform'])
    matriz_completa = [df_bbdd.columns.tolist()] + df_bbdd.fillna("").values.tolist()

    # Payload con ambas estructuras limpias
    payload = {
        "hoja_principal": matriz_completa,
        "hoja_som": matriz_som
    }

    print("🚀 Enviando datos a Google Sheets...")
    res1 = requests.post(WEBHOOK_URL1, json=payload)
    print("Sheet Flujo peatonal template:", "🎉 ¡Éxito!" if res1.text == "OK" else f"⚠️ Error: {res1.text}")
   
    # Llama a la función para ejecutar el proceso al correr el archivo
if __name__ == "__main__":
        ejecutar_y_enviar()
