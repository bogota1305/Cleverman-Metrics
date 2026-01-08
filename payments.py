import pandas as pd
import json
from modules.database_queries import execute_query
from modules.date_selector import open_date_selector
from modules.excel_creator import save_dataframe_to_excel, save_error_reasons_with_chart
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

def process_data(start_date, end_date):

    start_date = pd.to_datetime(start_date)
    query_start_date = (start_date - pd.DateOffset(months=1)).strftime('%Y-%m-%d')

    # Consulta para obtener los datos
    query_orders = f"""
    SELECT *
    FROM prod_sales_and_subscriptions.payments
    WHERE createdAt >= '{query_start_date} 00:00:00'
    AND createdAt < '{end_date} 00:00:00';
    """
    sP = execute_query(query_orders)

    # Convertir fechas
    sP['createdAt'] = pd.to_datetime(sP['createdAt'])
    sP['date'] = sP['createdAt'].dt.date

    # Agrupar por entityId para procesar métricas
    def calculate_metrics(group):
        group = group.sort_values('createdAt')

        # Identificar el primer pago fallido
        first_failed = group.iloc[0] if group.iloc[0]['status'] == 'FAILED' else None

        # Identificar el primer pago exitoso después de un fallo
        resolved_date = None
        if 'FAILED' in group['status'].values and 'SUCCESS' in group['status'].values:
            success_row = group[group['status'] == 'SUCCESS'].iloc[0]
            resolved_date = success_row['createdAt'].date()

        return pd.Series({
            'first_error_date': first_failed['createdAt'].date() if first_failed is not None else None,
            'resolved_date': resolved_date,
            'metadata': first_failed['metadata'] if first_failed is not None else None  # Guardar metadatos del primer error
        })

    grouped = sP.groupby('entityId').apply(calculate_metrics)

    # Filtrar grupos con errores válidos según el rango de fechas
    grouped['is_error_in_range'] = grouped['first_error_date'].apply(
        lambda x: 1 if x and x >= start_date.date() else 0
    )
    grouped['is_resolved_in_range'] = grouped['resolved_date'].apply(
        lambda x: 1 if x and x >= start_date.date() else 0
    )

    # Contar grupos válidos
    error_group_count = grouped['is_error_in_range'].sum()

    # Crear un DataFrame para métricas diarias de errores
    errors_by_day = (
        grouped[grouped['is_error_in_range'] == 1]
        .groupby('first_error_date')
        .size()
        .reset_index(name='daily_errors')
        .rename(columns={'first_error_date': 'date'})
    )

    # Crear un DataFrame para métricas diarias de resoluciones
    resolved_by_day = (
        grouped[grouped['is_resolved_in_range'] == 1]
        .groupby('resolved_date')
        .size()
        .reset_index(name='daily_resolved')
        .rename(columns={'resolved_date': 'date'})
    )

    # Combinar métricas diarias
    daily_summary = pd.merge(errors_by_day, resolved_by_day, on='date', how='outer').fillna(0)

    # Cálculo de totales generales
    total_errors = daily_summary['daily_errors'].sum()
    total_resolved = daily_summary['daily_resolved'].sum()

    total_errors_average = daily_summary['daily_errors'].mean()
    total_resolved_average = daily_summary['daily_resolved'].mean()

    # Resumen total
    totals_row = pd.DataFrame([{
        'date': 'Total',
        'daily_errors': total_errors,
        'daily_resolved': total_resolved
    }])

    # Concatenar resumen total con los datos diarios
    daily_summary = pd.concat([daily_summary, totals_row], ignore_index=True)

    # Calcular razones de error basadas en el primer fallo de cada grupo dentro del rango
    def extract_decline_code(metadata):
        try:
            metadata_dict = json.loads(metadata)
            return metadata_dict.get('stripeError', {}).get('error', {}).get('decline_code', 'unknown_error')
        except (json.JSONDecodeError, TypeError):
            return 'invalid_metadata'

    # Filtrar solo los primeros errores de cada grupo en el rango
    first_errors_in_range = grouped[grouped['is_error_in_range'] == 1]
    first_errors_in_range['decline_code'] = first_errors_in_range['metadata'].apply(extract_decline_code)

    # Contar razones de error
    error_reasons = first_errors_in_range['decline_code'].value_counts().reset_index()
    error_reasons.columns = ['decline_code', 'error_count']

    total_payments = [total_errors_average, total_resolved_average]

    return daily_summary, error_reasons, total_payments

def get_payments(start_date, end_date, folder_name, dropbox_var, drive_var):
    
    file_name = 'Payment Errors'
    daily_summary, error_reasons, total_payments = process_data(start_date, end_date)

    # Crear gráficos para resumen general
    columns_to_plot = ['daily_errors', 'daily_resolved']
    colors = ['#0000FF', '#008000']
    grafico_positions = ['H2', 'H24']

    # Guardar resumen general usando save_dataframe_to_excel
    save_dataframe_to_excel(folder_name, file_name, daily_summary, 'General', columns_to_plot, colors, grafico_positions, False)

    # Guardar razones de error en Excel con Pandas
    urls = save_error_reasons_with_chart(folder_name, file_name, error_reasons, True, dropbox_var, drive_var)

    return total_payments, urls
