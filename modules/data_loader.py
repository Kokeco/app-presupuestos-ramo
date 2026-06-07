import pandas as pd
import streamlit as st

@st.cache_data
def cargar_datos_budget(ruta_archivo):
    """
    Lee el archivo CSV de presupuesto detectando el separador automáticamente,
    limpia encabezados y transforma a formato largo.
    """
    try:
        # sep=None y engine='python' detectan automáticamente si es coma o punto y coma
        df = pd.read_csv(ruta_archivo, sep=None, engine='python', encoding='utf-8')
        
        # Limpiamos espacios en blanco invisibles en los nombres de las columnas
        df.columns = df.columns.str.strip()
        
        # Columnas descriptivas que NO son meses
        columnas_base = [
            'Resp', 'Desc Resp', 'VP', 'Gerencia', 'Proc', 
            'Desc Proc', 'Item', 'Desc Item', 'Classif', 'CC', 
            'FY26', 'FY27', 'FY28', 'FY29', 'FY30'
        ]
        
        # Filtramos solo las columnas base que realmente existan en el archivo
        columnas_base_reales = [col for col in columnas_base if col in df.columns]
        
        # Las columnas de meses serán todas las que NO estén en las columnas base
        columnas_meses = [col for col in df.columns if col not in columnas_base_reales]
        
        if 'Classif' not in columnas_base_reales:
            st.error("⚠️ No se encontró la columna 'Classif' en el archivo. Verifica los encabezados.")
            return None

        # Transformamos la estructura de formato ancho a largo
        df_largo = df.melt(
            id_vars=columnas_base_reales, 
            value_vars=columnas_meses,
            var_name='Periodo', 
            value_name='Monto'
        )
        
        # Limpieza de nulos y conversión estricta a números
        df_largo = df_largo.dropna(subset=['Monto'])
        
        # Si los números vienen como texto con formato especial, limpiamos caracteres extra
        if df_largo['Monto'].dtype == 'object':
            df_largo['Monto'] = df_largo['Monto'].astype(str).str.replace('.', '', regex=False)
            df_largo['Monto'] = df_largo['Monto'].str.replace(',', '.', regex=False)
            
        df_largo['Monto'] = pd.to_numeric(df_largo['Monto'], errors='coerce').fillna(0)
        
        # Limpiamos los datos internos de la columna de clasificación
        df_largo['Classif'] = df_largo['Classif'].astype(str).str.strip()
        df_largo['Escenario'] = '1. Budget Base'
        
        return df_largo

    except Exception as e:
        st.error(f"Error crítico al procesar la data operativa: {e}")
        return None
