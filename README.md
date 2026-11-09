# Automatización de Flujo Peatonal - Malls

Este repositorio contiene los scripts para la extracción, procesamiento y visualización del flujo peatonal de los centros comerciales de Parque Arauco.

---

## 🛠️ Estructura del Proyecto

* **`pipeline-templateflujo.py`**: 
  * Conecta a la base de datos en AWS Redshift.
  * Extrae los datos de flujo peatonal (`REM` y `SOM`) para los años 2025 y 2026.
  * Mapea y homologa propiedades, accesos y fechas.
  * Envía la información estructurada a Google Sheets a través de un Webhook.

* **`flujo-peatonal-codigo.js`**: 
  * Script de Google Apps Script vinculado a Google Sheets.
  * Procesa los datos recibidos desde Redshift.
  * Calcula variaciones porcentuales (vs. Año Anterior y vs. Presupuesto).
  * Genera el desglose diario de flujo, análisis por accesos y tablas de resumen.

---

## 🔄 Flujo de Trabajo (Workflow)
