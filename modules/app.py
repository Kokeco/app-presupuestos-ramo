import streamlit as st
import pandas as pd
import plotly.express as px
from modules.data_loader import cargar_datos_budget
from modules.forecast import calcular_forecast

st.set_page_config(page_title="App Control de Presupuestos", layout="wide")

st.title("📊 Panel Interactivo de Presupuestos y Forecast")
st.markdown("Herramienta de control de gestión para simulación de escenarios.")

ruta_archivo = "data/2026 - 2030.csv"
df_budget = cargar_datos_budget(ruta_archivo)

# Solo ejecutamos la lógica visual si la data se cargó y procesó correctamente
if df_budget is not None and not df_budget.empty:
    st.sidebar.header("⚙️ Ajustes de Escenario")
    
    var_combustible = st.sidebar.slider("Variación Consumo Combustible (%)", -20.0, 20.0, 5.0, 1.0)
    var_tasa_cambio = st.sidebar.slider("Variación Tasa de Cambio (%)", -15.0, 15.0, 3.0, 0.5)
    var_labor = st.sidebar.slider("Ajuste Costo Laboral (%)", -10.0, 10.0, 0.0, 1.0)
    
    df_forecast = calcular_forecast(df_budget, var_combustible, var_tasa_cambio, var_labor)
    df_consolidado = pd.concat([df_budget, df_forecast])
    
    total_budget = df_budget['Monto'].sum()
    total_forecast = df_forecast['Monto'].sum()
    diferencia = total_forecast - total_budget
    
    st.subheader("Resumen de Impacto Global")
    col1, col2, col3 = st.columns(3)
    col1.metric("Presupuesto Base (Budget)", f"${total_budget:,.0f}")
    col2.metric("Proyección (Forecast)", f"${total_forecast:,.0f}", f"${diferencia:,.0f} vs Base")
    
    st.markdown("---")
    
    st.subheader("Análisis Visual: Budget vs Forecast")
    resumen_grafico = df_consolidado.groupby(['Classif', 'Escenario'])['Monto'].sum().reset_index()
    
    fig = px.bar(
        resumen_grafico, 
        x='Classif', 
        y='Monto', 
        color='Escenario', 
        barmode='group',
        title="Comparativa de Costos por Categoría Operativa",
        labels={'Classif': 'Clasificación de Gasto', 'Monto': 'Monto ($)'},
        color_discrete_map={'1. Budget Base': '#1f77b4', '2. Forecast Proyectado': '#ff7f0e'}
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    with st.expander("Ver Base de Datos Transformada"):
        st.dataframe(df_forecast.head(100))
else:
    st.warning("⚠️ No se pudo procesar el archivo. Revisa que esté correctamente cargado en la carpeta 'data/'.")
