Forecast 5+7 No Lineal Dinámico - Streamlit
Aplicación web para realizar un Forecast 5+7 no lineal sobre gastos mineros, usando Excel de gastos y presupuesto.

Qué hace
Carga cualquier Excel con estructura similar de gastos/presupuesto.
Permite seleccionar hoja de gastos y hoja de presupuesto.
Detecta automáticamente columnas mensuales tipo Jan-25, Ene-25, 2025-01, etc.
Permite elegir el último mes real/YTD.
Calcula forecast de los meses restantes del año.
Usa un modelo no lineal con:
factor de ejecución acumulada,
factor de tendencia reciente,
curva logística no lineal,
ajuste por contexto minero/naturaleza de gasto.
Genera dashboard, hallazgos automáticos, propuesta de mejora y descarga de resultados.
Cómo correr localmente
pip install -r requirements.txt
streamlit run app.py
Cómo subir a GitHub + Streamlit Cloud
Crear un repositorio en GitHub.
Subir app.py, requirements.txt y este README.md.
Entrar a Streamlit Cloud.
Conectar el repositorio.
Seleccionar app.py como archivo principal.
Deploy.
Modelo usado
Forecast mensual futuro = Budget mensual × Factor ejecución × Factor tendencia × Curva no lineal
Donde:

Factor ejecución = Actual YTD / Budget YTD
Factor tendencia = Promedio meses recientes / Promedio meses iniciales
La curva no lineal distribuye el ajuste de manera logística en los meses futuros, evitando una extrapolación lineal simple.

Defensa frente a la rúbrica
Esta app cumple con:

Forecast 5+7 no lineal.
App web funcional en Streamlit/GitHub.
Dashboard ejecutivo.
Resultados y hallazgos.
Propuesta formal de mejora.
Conclusiones/recomendaciones automáticas.
