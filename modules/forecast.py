import pandas as pd

def calcular_forecast(df_base, var_combustible, var_tasa_cambio, var_labor):
    """
    Toma la base de datos limpia y aplica variaciones porcentuales de forma segura.
    """
    df_forecast = df_base.copy()
    
    # Pasamos temporalmente a minúsculas para asegurar que el filtro funcione siempre
    classif_lower = df_forecast['Classif'].str.lower()
    
    # 1. Ajuste al Combustible (Fuel)
    df_forecast.loc[classif_lower == 'fuel', 'Monto'] *= (1 + (var_combustible / 100))
    
    # 2. Ajuste por Tasa de Cambio (Expenses y Contractors)
    filtro_tasa = classif_lower.isin(['expenses', 'contractors'])
    df_forecast.loc[filtro_tasa, 'Monto'] *= (1 + (var_tasa_cambio / 100))
    
    # 3. Ajuste al Costo Laboral (Labor)
    df_forecast.loc[classif_lower == 'labor', 'Monto'] *= (1 + (var_labor / 100))
    
    df_forecast['Escenario'] = '2. Forecast Proyectado'
    
    return df_forecast
