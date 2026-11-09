import os
import psycopg2
import requests
from dotenv import load_dotenv

# Apunta directamente al .env guardado dentro de .venv
load_dotenv(".venv/.env", override=True)

# Conexión a Redshift
conn = psycopg2.connect(
    dbname=os.getenv("REDSHIFT_DBNAME"),
    user=os.getenv("REDSHIFT_USER"),
    password=os.getenv("REDSHIFT_PASSWORD"),
    host=os.getenv("REDSHIFT_HOST"),
    port=os.getenv("REDSHIFT_PORT"),
    connect_timeout=10
)
cursor = conn.cursor()

print("--- Extrayendo datos 2026 ---")

# 1. QUERY FLUJO ACCESO (Detalle por entrada)
query_acceso = """
SELECT
    CASE id_property
        WHEN 'MAP' THEN 'ASA'
        WHEN 'PLZ' THEN 'MAE'
        WHEN 'MQU' THEN 'MAQ'
        WHEN 'BAV' THEN 'BOM'
        ELSE id_property
    END AS mall
    ,date::text AS dia
    ,LAST_VALUE(entrance_name) OVER(
        PARTITION BY id_property, entrance_key
        ORDER BY date ASC
        ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
    ) AS acceso
    ,SUM(in_reported) AS flujo_real
FROM parauco_analysis_opscentroscomerciales.agg_entsal_footfall_flow_cl
WHERE
    (id_platform = 'REM' OR id_platform = 'SOM')
    AND date >= '2026-11-01' # cambiar según la fecha de inicio del flujo que se quiera extraer
    AND date <= CURRENT_DATE
GROUP BY ALL
ORDER BY mall, date ASC, acceso;
"""

cursor.execute(query_acceso)
rows_acceso = cursor.fetchall()

# Diccionario de equivalencias para caracteres especiales/corruptos del histórico 2025
REEMPLAZOS_CARACTERES_ESPECIALES = {
    "Nivel 1 - Escala MecÃ¡nica Plaza General": "Nivel 1 - Escala Mecánica Plaza General",
    "Acceso 01 Principal Lado PillÃ¬n": "Acceso 01 Principal Lado Pillín",
    "Acceso de PÃºblico": "Acceso de Público"
}

filas_acceso_limpias = []
for row in rows_acceso:
    mall, dia, acceso, flujo_real = row
    
    if acceso:
        # Step 1: Elimina todos los espacios extras al final/inicio
        acceso = acceso.strip()
        
        # Step 2: Mapea a la cadena con caracteres especiales del 2025 si corresponde
        if acceso in REEMPLAZOS_CARACTERES_ESPECIALES:
            acceso = REEMPLAZOS_CARACTERES_ESPECIALES[acceso]
            
    filas_acceso_limpias.append([mall, dia, acceso, flujo_real])

data_acceso = [["MALL", "DÍA", "ACCESO", "FLUJO REAL"]] + filas_acceso_limpias

# 2. QUERY FLUJO PPTO (Total por Mall)
query_ppto = """
SELECT
    CASE id_property
        WHEN 'MAP' THEN 'ASA'
        WHEN 'PLZ' THEN 'MAE'
        WHEN 'MQU' THEN 'MAQ'
        WHEN 'BAV' THEN 'BOM'
        ELSE id_property
    END AS mall
    ,date::text AS dia
    ,SUM(in_reported) AS flujo_real
FROM parauco_analysis_opscentroscomerciales.agg_entsal_footfall_flow_cl
WHERE
    (id_platform = 'REM' OR id_platform = 'SOM')
    AND date >= '2026-11-01' # cambiar según la fecha de inicio del flujo que se quiera extraer
    AND date <= CURRENT_DATE
GROUP BY ALL
ORDER BY date ASC, mall;
"""

cursor.execute(query_ppto)
rows_ppto = cursor.fetchall()
data_ppto = [["MALL", "DÍA", "FLUJO REAL"]] + [list(r) for r in rows_ppto]

cursor.close()
conn.close()

# 3. ENVIAR AL NUEVO GOOGLE SHEET
# Lee la nueva URL configurada en el archivo .env
webhook_url_looker = os.getenv("WEBHOOK_URL_LOOKER")

# Enviar Pestaña 1 (Acceso)
req1 = requests.post(webhook_url_looker, json={"sheet_name": "extraer_flujo_acceso", "data": data_acceso})
print("Pestaña 'extraer_flujo_acceso':", req1.status_code)

# Enviar Pestaña 2 (PPTO)
req2 = requests.post(webhook_url_looker, json={"sheet_name": "extraer_flujo_ppto", "data": data_ppto})
print("Pestaña 'extraer_flujo_ppto':", req2.status_code)

print("¡Datos enviados con éxito al nuevo Google Sheet!🎉")
