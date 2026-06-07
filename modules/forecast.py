import pandas as pd

def calcular_forecast(df_base, var_combustible, var_tasa_cambio, var_labor):
    """
    Recibe el DataFrame base y aplica las variaciones porcentuales del usuario.
    """
    df_forecast = df_base.copy()
    
    # 1. Aplicar variación al Combustible (Fuel)
    filtro_fuel = df_forecast['Classif'] == 'Fuel'
    df_forecast.loc[filtro_fuel, 'Monto'] = df_forecast.loc[filtro_fuel, 'Monto'] * (1 + (var_combustible / 100))
    
    # 2. Aplicar Tasa de Cambio (Afecta a Expenses y Contractors)
    filtro_tasa = df_forecast['Classif'].isin(['Expenses', 'Contractors'])
    df_forecast.loc[filtro_tasa, 'Monto'] = df_forecast.loc[filtro_tasa, 'Monto'] * (1 + (var_tasa_cambio / 100))
    
    # 3. Aplicar variación a Mano de Obra (Labor)
    filtro_labor = df_forecast['Classif'] == 'Labor'
    df_forecast.loc[filtro_labor, 'Monto'] = df_forecast.loc[filtro_labor, 'Monto'] * (1 + (var_labor / 100))
    
    # Cambiamos la etiqueta para diferenciarlo en los gráficos
    df_forecast['Escenario'] = '2. Forecast Proyectado'
    
    return df_forecast
