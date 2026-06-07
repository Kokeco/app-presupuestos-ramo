import pandas as pd
import streamlit as st

@st.cache_data
def cargar_datos_budget(ruta_archivo):
    """
    Lee el archivo CSV de presupuesto, lo limpia y lo transforma a formato largo.
    """
    try:
        # Cargamos el archivo CSV
        df = pd.read_csv(ruta_archivo)
        
        # Columnas descriptivas que NO son meses
        columnas_base = [
            'Resp', 'Desc Resp', 'VP', 'Gerencia', 'Proc', 
            'Desc Proc', 'Item', 'Desc Item', 'Classif', 'CC', 
            'FY26', 'FY27', 'FY28', 'FY29', 'FY30' # Ajusta estos FY según tu archivo
        ]
        
        # Filtramos para asegurarnos de que las columnas base existan en el archivo
        columnas_base_reales = [col for col in columnas_base if col in df.columns]
        
        # Identificamos las columnas que son estrictamente meses (Jan-26, Feb-26, etc.)
        columnas_meses = [col for col in df.columns if col not in columnas_base_reales]
        
        # Unpivot: pasamos de formato ancho a largo
        df_largo = df.melt(
            id_vars=columnas_base_reales, 
            value_vars=columnas_meses,
            var_name='Periodo', 
            value_name='Monto'
        )
        
        # Limpieza: eliminamos nulos y aseguramos que el monto sea numérico
        df_largo = df_largo.dropna(subset=['Monto'])
        df_largo['Monto'] = pd.to_numeric(df_largo['Monto'], errors='coerce').fillna(0)
        
        # Limpiamos espacios en blanco en la clasificación
        df_largo['Classif'] = df_largo['Classif'].str.strip()
        df_largo['Escenario'] = '1. Budget Base'
        
        return df_largo

    except Exception as e:
        st.error(f"Error al cargar la data. Verifica que el archivo exista: {e}")
        return None
