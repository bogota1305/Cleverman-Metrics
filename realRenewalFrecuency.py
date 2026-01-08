import os
import pandas as pd
import numpy as np
from modules.database_queries import execute_query
from datetime import datetime
from uploadCloud import upload_to_drive, upload_to_dropbox

def renewalFrequency(query, fileName):
    # Obtener datos desde la base de datos
    data = execute_query(query)
    
    # Convertir a DataFrame
    df = pd.DataFrame(data, columns=['order_number', 'subscription_id', 'created_at', 'legacy_category', 'delivery_frequency'])
    
    # Convertir fechas
    df['created_at'] = pd.to_datetime(df['created_at'])
    
    # Función para extraer frecuencia en semanas desde delivery_frequency
    def extract_frequency(delivery_freq):
        try:
            if pd.isna(delivery_freq):
                return None
            days = int(delivery_freq.split()[0])
            return days / 7  # Convertir a semanas
        except:
            return None
    
    df['frequency_weeks'] = df['delivery_frequency'].apply(extract_frequency)
    
    # Procesar cada suscripción
    results = []
    for sub_id, group in df.groupby('subscription_id'):
        if len(group) >= 2:
            # Ordenar por fecha (más reciente primero)
            sorted_orders = group.sort_values('created_at', ascending=False)
            
            # Tomar las 2 órdenes más recientes
            last_two = sorted_orders.head(2)
            
            # Calcular diferencia en días
            days_diff = (last_two.iloc[0]['created_at'] - last_two.iloc[1]['created_at']).days
            weeks_diff = days_diff / 7
            
            results.append({
                'subscription_id': sub_id,
                'days_between_orders': days_diff,
                'weeks_between_orders': weeks_diff,
                'frequency_weeks': last_two.iloc[0]['frequency_weeks'],
                'category': last_two.iloc[0]['legacy_category']
            })
    
    if not results:
        print(f"No hay suficientes datos para generar reporte ({fileName})")
        return
    
    result_df = pd.DataFrame(results)
    
    # Calcular promedios generales
    general_avg_weeks = result_df['weeks_between_orders'].mean()
    general_avg_days = result_df['days_between_orders'].mean()
    freq_avg_weeks = result_df['frequency_weeks'].mean()
    
   # Calcular modas generales
    general_mode_weeks = result_df['weeks_between_orders'].mode()
    general_mode_days = result_df['days_between_orders'].mode()
    freq_mode_weeks = result_df['frequency_weeks'].mode()
    
    # Calcular promedios y modas por categoría
    beard_data = result_df[result_df['category'] == 'BEARD']
    beard_avg_weeks = beard_data['weeks_between_orders'].mean() if not beard_data.empty else 0
    beard_avg_days = beard_data['days_between_orders'].mean() if not beard_data.empty else 0
    beard_mode_weeks = beard_data['weeks_between_orders'].mode() if not beard_data.empty else pd.Series([0])
    beard_mode_days = beard_data['days_between_orders'].mode() if not beard_data.empty else pd.Series([0])
    
    hair_data = result_df[result_df['category'] == 'HAIR']
    hair_avg_weeks = hair_data['weeks_between_orders'].mean() if not hair_data.empty else 0
    hair_avg_days = hair_data['days_between_orders'].mean() if not hair_data.empty else 0
    hair_mode_weeks = hair_data['weeks_between_orders'].mode() if not hair_data.empty else pd.Series([0])
    hair_mode_days = hair_data['days_between_orders'].mode() if not hair_data.empty else pd.Series([0])
    
    # Función para formatear la moda (puede haber múltiples valores)
    def format_mode(mode_series):
        if len(mode_series) == 0:
            return "No data"
        elif len(mode_series) == 1:
            return f"{mode_series.iloc[0]:.2f}"
        else:
            return ", ".join([f"{val:.2f}" for val in mode_series.head(3)])  # Máximo 3 valores
    
    # Crear DataFrames para las hojas de Excel
    details_df = result_df[[
        'subscription_id', 
        'days_between_orders', 
        'weeks_between_orders',
        'frequency_weeks',
        'category'
    ]]
    
    # DataFrame de comparación expandido con modas
    comparison_df = pd.DataFrame({
        'Metric': [
            'Configured Frequency (avg)',
            'Configured Frequency (mode)',
            'Actual Frequency (avg)',
            'Actual Frequency (mode)',
            'BEARD Frequency (avg)',
            'BEARD Frequency (mode)',
            'HAIR Frequency (avg)',
            'HAIR Frequency (mode)'
        ],
        'Weeks': [
            freq_avg_weeks,
            format_mode(freq_mode_weeks),
            general_avg_weeks,
            format_mode(general_mode_weeks),
            beard_avg_weeks,
            format_mode(beard_mode_weeks),
            hair_avg_weeks,
            format_mode(hair_mode_weeks)
        ],
        'Days': [
            freq_avg_weeks * 7,
            format_mode(freq_mode_weeks * 7) if not freq_mode_weeks.empty else "No data",
            general_avg_days,
            format_mode(general_mode_days),
            beard_avg_days,
            format_mode(beard_mode_days),
            hair_avg_days,
            format_mode(hair_mode_days)
        ]
    })
    
    # Crear archivo Excel
    with pd.ExcelWriter(fileName) as writer:
        # Hoja de detalles
        details_df.to_excel(writer, sheet_name='Subscription Details', index=False)
        
        # Hoja de comparación
        comparison_df.to_excel(writer, sheet_name='Frequency Comparison', index=False)
        
        # Ajustar formatos automáticamente al contenido
        workbook = writer.book
        
        # Función para ajustar ancho de columnas
        def auto_adjust_columns(worksheet, df):
            for idx, col in enumerate(df.columns):
                max_len = max((
                    df[col].astype(str).map(len).max(),  # Longitud máxima de los datos
                    len(str(col))  # Longitud del nombre de la columna
                )) + 2  # Pequeño margen
                worksheet.set_column(idx, idx, max_len)
        
        # Aplicar a ambas hojas
        auto_adjust_columns(writer.sheets['Subscription Details'], details_df)
        auto_adjust_columns(writer.sheets['Frequency Comparison'], comparison_df)
    
    print(f"Reporte generado: {fileName}")
    print(f"Suscripciones analizadas: {len(result_df)}")
    print(f"Frecuencia promedio real: {general_avg_weeks:.2f} semanas")

def realRenewalFrequency(start_date, end_date, folder_name):
    # Consulta SQL para suscripciones sin Full Control
    queryNoFC = f"""
    SELECT
        fo.order_number,
        fo.subscription_id,    
        fo.created_at,
        sv.legacy_category,
        sv.delivery_frequency
    FROM 
        bi.fact_orders fo
    JOIN
        prod_sales_and_subscriptions.subscriptions su ON fo.subscription_id = su.id
    JOIN
        prod_sales_and_subscriptions.subscriptions_view sv ON fo.subscription_id = sv.subscription_id
    WHERE 
        fo.subscription_id IS NOT NULL
        AND (
            (
                su.additionalFields ->> "$.sms_renewal" = "false" 
                OR su.additionalFields ->> "$.sms_renewal" IS NULL
            ) 
        )
        AND su.status != 'CANCELLED'
        AND fo.status != 'CANCELLED'
        AND su.createdAt > '{start_date}' AND su.createdAt < '{end_date}'
    """

    # Consulta SQL para suscripciones con Full Control
    queryFC = f"""
    SELECT
        fo.order_number,
        fo.subscription_id,    
        fo.created_at,
        sv.legacy_category,
        sv.delivery_frequency
    FROM 
        bi.fact_orders fo
    JOIN
        prod_sales_and_subscriptions.subscriptions su ON fo.subscription_id = su.id
    JOIN
        prod_sales_and_subscriptions.subscriptions_view sv ON fo.subscription_id = sv.subscription_id
    WHERE 
        fo.subscription_id IS NOT NULL
        AND su.additionalFields ->> "$.sms_renewal" = "true" 
        AND su.status != 'CANCELLED'
        AND fo.status != 'CANCELLED'
        AND su.createdAt > '{start_date}' AND su.createdAt < '{end_date}'
    """

    # Consulta SQL para todas las suscripciones
    queryAll = f"""
    SELECT
        fo.order_number,
        fo.subscription_id,    
        fo.created_at,
        sv.legacy_category,
        sv.delivery_frequency
    FROM 
        bi.fact_orders fo
    JOIN
        prod_sales_and_subscriptions.subscriptions su ON fo.subscription_id = su.id
    JOIN
        prod_sales_and_subscriptions.subscriptions_view sv ON fo.subscription_id = sv.subscription_id
    WHERE 
        fo.subscription_id IS NOT NULL    
        AND su.status != 'CANCELLED'
        AND fo.status != 'CANCELLED'
        AND su.createdAt > '{start_date}' AND su.createdAt < '{end_date}'
    """

    saveFile(folder_name, 'renewal_frequency_noFC.xlsx', queryNoFC)
    saveFile(folder_name, 'renewal_frequency_FC.xlsx', queryFC)
    saveFile(folder_name, 'renewal_frequency_ALL.xlsx', queryAll)

def saveFile(folder_name, file_name, query):

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    # Construir la ruta completa del archivo
    full_path = os.path.join(folder_name, file_name)
    renewalFrequency(query, full_path)
    #upload_to_drive(full_path, folder_id="1F1VZxlp5IxkQEo4WD0Bt8VEJZ28OhGut")
    #upload_to_dropbox(full_path, dropbox_path=f"/MyReports/{folder_name}/{file_name}")

# realRenewalFrequency('2024-01-01', '2025-01-01', 'renewal')
    