import pandas as pd
import json
from datetime import datetime, timedelta

from modules.database_queries import execute_query
from modules.excel_creator import save_dataframe_to_excel

def get_expected_renewals(start_date, end_date, folder_name):
    # Retroceder 8 semanas desde la fecha de inicio
    start_date_dt = datetime.strptime(start_date, '%Y-%m-%d')
    end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')
    adjusted_start_date = start_date_dt - timedelta(weeks=8)

    # Consulta SQL para obtener los datos de subscriptions
    query_subscriptions = f"""
    SELECT createdAt, updatedAt, additionalFields, status
    FROM sales_and_subscriptions.subscriptions
    WHERE updatedAt >= '{adjusted_start_date.strftime('%Y-%m-%d')} 00:00:00'
      AND updatedAt < '{end_date} 00:00:00'
      AND status IN ('ACTIVE') ;
    """
    
    # Ejecutar la consulta (la función execute_query está definida en otro archivo)
    subscriptions = execute_query(query_subscriptions)

    # Asegurar que los datos están en un DataFrame
    subscriptions = pd.DataFrame(subscriptions, columns=['updatedAt', 'additionalFields'])

    # Convertir createdAt a tipo datetime
    subscriptions['updatedAt'] = pd.to_datetime(subscriptions['updatedAt'])

    # Procesar el campo additionalFields para obtener el valor 'frequency'
    def extract_frequency(additional_fields):
        try:
            data = json.loads(additional_fields)
            frequency = data.get('frequency', '')
            if frequency.startswith('Every'):
                weeks = int(frequency.split()[1])
                return weeks
        except (json.JSONDecodeError, ValueError, IndexError):
            return None

    subscriptions['frequency_weeks'] = subscriptions['additionalFields'].apply(extract_frequency)

    # Filtrar filas con frecuencia válida
    subscriptions = subscriptions.dropna(subset=['frequency_weeks'])

    # Calcular el día esperado de renovación
    subscriptions['expected_renewal'] = subscriptions['updatedAt'] + subscriptions['frequency_weeks'].apply(lambda x: timedelta(weeks=x))

    # Filtrar renovaciones dentro del rango de fechas proporcionado
    subscriptions = subscriptions[(subscriptions['expected_renewal'] >= start_date_dt) &
                                  (subscriptions['expected_renewal'] <= end_date_dt)]

    # Contar renovaciones por día esperado
    expected_renewals = subscriptions.groupby(subscriptions['expected_renewal'].dt.date).size().reset_index(name='renewals_count')

    # Renombrar columnas para claridad
    expected_renewals.rename(columns={'expected_renewal': 'date'}, inplace=True)

    # Consulta SQL para obtener los datos de fact_orders
    query_orders = f"""
    SELECT created_at, order_plan, recurrent, units
    FROM bi.fact_orders
    WHERE created_at >= '{start_date} 00:00:00'
      AND created_at < '{end_date} 00:00:00';
    """

    # Ejecutar la consulta
    fact_orders = execute_query(query_orders)

    # Asegurar que los datos están en un DataFrame
    fact_orders = pd.DataFrame(fact_orders, columns=['created_at', 'order_plan', 'recurrent', 'units'])

    # Filtrar los datos según las condiciones especificadasunits
    fact_orders = fact_orders[(fact_orders['order_plan'].isin(['SUBSCRIPTION'])) & (fact_orders['recurrent'] == 1) & (fact_orders['units'] == 1)]

    # Convertir createdAt a tipo datetime
    fact_orders['created_at'] = pd.to_datetime(fact_orders['created_at'])

    # Contar renovaciones obtenidas por día
    fact_orders['date'] = fact_orders['created_at'].dt.date
    obtained_renewals = fact_orders.groupby('date').size().reset_index(name='obtained_renewals')

    # Combinar los datos de renovaciones esperadas y obtenidas
    combined_renewals = pd.merge(expected_renewals, obtained_renewals, on='date', how='outer').fillna(0)

    # Asegurar tipos de datos correctos después del merge
    combined_renewals['renewals_count'] = combined_renewals['renewals_count'].astype(int)
    combined_renewals['obtained_renewals'] = combined_renewals['obtained_renewals'].astype(int)
    
     # Calcular totales
    total_row = pd.DataFrame({
        'date': ['Total'],
        'renewals_count': [combined_renewals['renewals_count'].sum()],
        'obtained_renewals': [combined_renewals['obtained_renewals'].sum()]
    })

    # Agregar fila de totales
    combined_renewals = pd.concat([combined_renewals, total_row], ignore_index=True)

    # Calcular porcentajes
    combined_renewals['percentage'] = combined_renewals.apply(
        lambda row: float((row['obtained_renewals'] / row['renewals_count'])) \
        if row['renewals_count'] > 0 else 0,
        axis=1
    )

    file_name = 'Renewals'

    # Crear gráficos para resumen general
    columns_to_plot = ['renewals_count', 'obtained_renewals', 'percentage']
    colors = ['#0000FF', '#008000', '#FF0000']
    grafico_positions = ['H2', 'H24', 'H46']
        
    save_dataframe_to_excel(folder_name, file_name, combined_renewals, 'General', columns_to_plot, colors, grafico_positions)

    return combined_renewals
