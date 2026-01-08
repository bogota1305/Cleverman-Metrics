import os
import pandas as pd
import numpy as np
import json
import re
from modules.database_queries import execute_query
from datetime import datetime, timedelta
from uploadCloud import upload_to_drive, upload_to_dropbox

def extract_weeks_from_frequency(freq_str):
    """Extrae el número de semanas de una cadena de frecuencia"""
    try:
        if 'Every' in freq_str:
            # Buscar patrones como "Every 3 weeks" o "Every 4 weeks (RECO)"
            match = re.search(r'Every\s+(\d+)\s+weeks', freq_str)
            if match:
                return int(match.group(1))
        return None
    except:
        return None

def calculate_frequency_change_difference(frequencies_dict):
    """Calcula la diferencia entre los dos últimos cambios de frecuencia válidos"""
    try:
        if not frequencies_dict:
            return None
        
        # Convertir a lista de tuplas (frecuencia, fecha) y ordenar por fecha, filtrando solo cambios válidos
        freq_changes = []
        for freq, date_str in frequencies_dict.items():
            weeks = extract_weeks_from_frequency(freq)
            if weeks is not None:  # Solo incluir cambios con formato "Every X weeks"
                try:
                    date = datetime.strptime(date_str.split('.')[0], '%Y-%m-%d %H:%M:%S')
                    freq_changes.append((weeks, date))
                except:
                    continue
        
        # Ordenar por fecha (más reciente primero)
        freq_changes.sort(key=lambda x: x[1], reverse=True)
        
        # Tomar las dos frecuencias más recientes que sean válidas
        valid_changes = [fc for fc in freq_changes if fc[0] is not None]
        if len(valid_changes) >= 2:
            recent_weeks = valid_changes[0][0]
            previous_weeks = valid_changes[1][0]
            return recent_weeks - previous_weeks
        
        return None
    except:
        return None

def renewalFrequency(query_main, fileName):
    # Obtener datos desde la base de datos para el query principal
    print("Ejecutando query principal...")
    data_main = execute_query(query_main)
    print(f"Query principal retornó {len(data_main)} registros")
    
    # Query para payment errors
    query_payment_errors = """
    WITH last_orders AS (
    SELECT
        fo.subscription_id,
        fo.created_at,
        ROW_NUMBER() OVER (PARTITION BY fo.subscription_id ORDER BY fo.created_at DESC) AS rn
    FROM bi.fact_orders fo
    WHERE fo.status != 'CANCELLED'
    ),
    order_ranges AS (
        SELECT
            lo.subscription_id,
            MAX(CASE WHEN lo.rn = 1 THEN lo.created_at END) AS last_order_date,
            MAX(CASE WHEN lo.rn = 2 THEN lo.created_at END) AS second_last_order_date
        FROM last_orders lo
        GROUP BY lo.subscription_id
    ),
    payment_errors AS (
        SELECT 
            r.subscriptionId,
            GREATEST(COUNT(*) - 1, 0) AS error_count
        FROM prod_sales_and_subscriptions.payments p
        JOIN prod_sales_and_subscriptions.renewals r 
            ON p.entityId = r.id
        JOIN order_ranges o
            ON r.subscriptionId = o.subscription_id
        AND p.createdAt BETWEEN o.second_last_order_date AND o.last_order_date
        WHERE p.entityId LIKE 'RE%'
        GROUP BY r.subscriptionId, p.entityId
    )
    SELECT
        subscriptionId,
        SUM(error_count) AS payment_errors
    FROM payment_errors
    GROUP BY subscriptionId
    ORDER BY payment_errors DESC;

    """
    
    # Query para cambios de frecuencia
    query_frequency_changes = """
    WITH base AS (
        SELECT
            re.subscriptionId,
            fe.createdAt,
            TRIM(BOTH ')' FROM SUBSTRING_INDEX(SUBSTRING_INDEX(fe.description, '(', -1), ')', 1)) AS frequency
        FROM prod_sales_and_subscriptions.fees fe
        JOIN prod_sales_and_subscriptions.renewals re 
            ON fe.entityId = re.id
        WHERE fe.description LIKE 'Sub%'
        AND fe.entityId LIKE 'RE%'
    ),
    with_lag AS (
        SELECT
            subscriptionId,
            createdAt,
            frequency,
            LAG(frequency) OVER (PARTITION BY subscriptionId ORDER BY createdAt) AS prev_frequency
        FROM base
    ),
    only_changes AS (
        SELECT
            subscriptionId,
            createdAt,
            frequency
        FROM with_lag
        WHERE frequency <> prev_frequency OR prev_frequency IS NULL
    ),
    final AS (
        SELECT
            subscriptionId,
            CONCAT(
                '{',
                GROUP_CONCAT(
                    CONCAT('"', frequency, '": "', createdAt, '"')
                    ORDER BY createdAt
                    SEPARATOR ','
                ),
                '}'
            ) AS frequency_changes_json,
            COUNT(DISTINCT frequency) AS num_changes
        FROM only_changes
        GROUP BY subscriptionId
    )
    SELECT
        subscriptionId,
        frequency_changes_json
    FROM final
    WHERE num_changes > 1;  
    """
    
    # Ejecutar queries adicionales
    print("Ejecutando query de payment errors...")
    data_payment_errors = execute_query(query_payment_errors)
    print(f"Query payment errors retornó {len(data_payment_errors)} registros")
    
    print("Ejecutando query de cambios de frecuencia...")
    data_frequency_changes = execute_query(query_frequency_changes)
    print(f"Query cambios de frecuencia retornó {len(data_frequency_changes)} registros")
    
    # Convertir a DataFrames
    df_main = pd.DataFrame(data_main, columns=[
        'subscription_id', 
        'legacy_category', 
        'delivery_frequency',
        'snooze',
        'last_order_date',
        'second_last_order_date',
        'days_diff'
    ])
    
    df_payment_errors = pd.DataFrame(data_payment_errors, columns=['subscriptionId', 'payment_errors'])
    
    # Convertir frequency changes
    if isinstance(data_frequency_changes, pd.DataFrame):
        df_frequency_changes = data_frequency_changes
        if not df_frequency_changes.empty and len(df_frequency_changes.columns) >= 2:
            df_frequency_changes.columns = ['subscriptionId', 'frequency_changes_json']
        else:
            df_frequency_changes = pd.DataFrame(columns=['subscriptionId', 'frequency_changes_json'])
    elif isinstance(data_frequency_changes, list) and len(data_frequency_changes) > 0:
        if len(data_frequency_changes[0]) >= 2:
            df_frequency_changes = pd.DataFrame(data_frequency_changes, columns=['subscriptionId', 'frequency_changes_json'])
        else:
            df_frequency_changes = pd.DataFrame(columns=['subscriptionId', 'frequency_changes_json'])
    else:
        df_frequency_changes = pd.DataFrame(columns=['subscriptionId', 'frequency_changes_json'])
    
    # Convertir fechas en df_main
    df_main['last_order_date'] = pd.to_datetime(df_main['last_order_date'])
    df_main['second_last_order_date'] = pd.to_datetime(df_main['second_last_order_date'])
    
    # Asegurarnos de que los tipos de datos sean consistentes
    df_main['subscription_id'] = df_main['subscription_id'].astype(str)
    if not df_frequency_changes.empty:
        df_frequency_changes['subscriptionId'] = df_frequency_changes['subscriptionId'].astype(str)
    
    # Unir los datos
    # Unir payment errors
    df_payment_errors['subscriptionId'] = df_payment_errors['subscriptionId'].astype(str)
    df_main = df_main.merge(
        df_payment_errors, 
        left_on='subscription_id', 
        right_on='subscriptionId', 
        how='left'
    ).drop('subscriptionId', axis=1)
    
    # Unir frequency changes solo si hay datos
    if not df_frequency_changes.empty:
        df_main = df_main.merge(
            df_frequency_changes, 
            left_on='subscription_id', 
            right_on='subscriptionId', 
            how='left'
        ).drop('subscriptionId', axis=1)
    else:
        df_main['frequency_changes_json'] = '{}'
    
    # Llenar NaN values
    df_main['payment_errors'] = df_main['payment_errors'].fillna(0)
    df_main['frequency_changes_json'] = df_main['frequency_changes_json'].fillna('{}')
    
    # ANALIZAR CAMBIOS DE FRECUENCIA VÁLIDOS (entre penúltima y última orden)
    frequency_changes_data = []  # True/False si tiene cambio válido en el período
    frequency_change_differences = []  # Diferencia en semanas entre los dos últimos cambios válidos
    
    print("Analizando cambios de frecuencia válidos...")
    valid_changes_count = 0
    
    for idx, row in df_main.iterrows():
        has_valid_frequency_change = False
        frequency_change_diff = None
        
        try:
            freq_json = row['frequency_changes_json']
            if freq_json and freq_json != '{}':
                frequencies_dict = json.loads(freq_json)
                
                if frequencies_dict:
                    # Filtrar solo cambios con formato "Every X weeks"
                    valid_changes = {}
                    for freq, date_str in frequencies_dict.items():
                        weeks = extract_weeks_from_frequency(freq)
                        if weeks is not None:  # Solo incluir cambios válidos
                            try:
                                change_date = datetime.strptime(date_str.split('.')[0], '%Y-%m-%d %H:%M:%S')
                                valid_changes[freq] = date_str
                                
                                # Verificar si este cambio está entre la penúltima y última orden
                                if (pd.notna(row['second_last_order_date']) and 
                                    pd.notna(row['last_order_date']) and
                                    row['second_last_order_date'] <= change_date <= row['last_order_date']):
                                    has_valid_frequency_change = True
                                
                            except:
                                continue
                    
                    # Si encontramos cambios válidos en el período
                    if has_valid_frequency_change:
                        valid_changes_count += 1
                        
                        # Calcular diferencia entre los dos últimos cambios válidos
                        frequency_change_diff = calculate_frequency_change_difference(valid_changes)
                        if frequency_change_diff is not None:
                            frequency_change_differences.append(frequency_change_diff)
                
        except (json.JSONDecodeError, ValueError):
            pass
        
        frequency_changes_data.append(has_valid_frequency_change)
    
    print(f"Total de cambios de frecuencia válidos en el período: {valid_changes_count}")
    
    # Agregar columna de cambios de frecuencia válidos (ELIMINAMOS days_since_last_frequency_change)
    df_main['has_valid_frequency_change'] = frequency_changes_data
    
    # Convertir snooze a numérico (1 = snooze, 0 = no snooze)
    df_main['snooze'] = df_main['snooze'].apply(lambda x: 1 if str(x).strip() == '1' else 0)
    
    # Calcular semanas desde días
    df_main['weeks_diff'] = df_main['days_diff'] / 7
    
    # Función para extraer frecuencia configurada en semanas
    def extract_frequency(delivery_freq):
        try:
            if pd.isna(delivery_freq):
                return None
            days = int(delivery_freq.split()[0])
            return days / 7  # Convertir a semanas
        except:
            return None
    
    df_main['frequency_weeks'] = df_main['delivery_frequency'].apply(extract_frequency)
    
    # Filtrar suscripciones con datos válidos (days_diff no nulo)
    result_df = df_main[df_main['days_diff'].notna()].copy()
    
    if len(result_df) == 0:
        print(f"No hay suficientes datos para generar reporte ({fileName})")
        return
    
    # ANÁLISIS DE SNOOZE Y PAYMENT ERRORS
    total_subscriptions = len(result_df)
    total_with_snooze = result_df['snooze'].sum()
    total_with_payment_errors = (result_df['payment_errors'] > 0).sum()
    total_with_valid_frequency_changes = result_df['has_valid_frequency_change'].sum()
    
    print(f"RESUMEN FINAL:")
    print(f"Total suscripciones: {total_subscriptions}")
    print(f"Con snooze: {total_with_snooze}")
    print(f"Con payment errors: {total_with_payment_errors}")
    print(f"Con cambios de frecuencia válidos en el período: {total_with_valid_frequency_changes}")
    
    # Calcular promedio de diferencias de frecuencia
    avg_frequency_change = np.mean(frequency_change_differences) if frequency_change_differences else 0
    
    # Calcular porcentajes
    snooze_percentage = (total_with_snooze / total_subscriptions * 100) if total_subscriptions > 0 else 0
    payment_errors_percentage = (total_with_payment_errors / total_subscriptions * 100) if total_subscriptions > 0 else 0
    frequency_changes_percentage = (total_with_valid_frequency_changes / total_subscriptions * 100) if total_subscriptions > 0 else 0
    
    # Calcular promedios generales
    general_avg_weeks = result_df['weeks_diff'].mean()
    general_avg_days = result_df['days_diff'].mean()
    freq_avg_weeks = result_df['frequency_weeks'].mean()
    
    # Calcular modas generales
    general_mode_weeks = result_df['weeks_diff'].mode()
    general_mode_days = result_df['days_diff'].mode()
    freq_mode_weeks = result_df['frequency_weeks'].mode()
    
    # Calcular promedios y modas por categoría
    beard_data = result_df[result_df['legacy_category'] == 'BEARD']
    beard_avg_weeks = beard_data['weeks_diff'].mean() if not beard_data.empty else 0
    beard_avg_days = beard_data['days_diff'].mean() if not beard_data.empty else 0
    beard_mode_weeks = beard_data['weeks_diff'].mode() if not beard_data.empty else pd.Series([0])
    beard_mode_days = beard_data['days_diff'].mode() if not beard_data.empty else pd.Series([0])
    
    hair_data = result_df[result_df['legacy_category'] == 'HAIR']
    hair_avg_weeks = hair_data['weeks_diff'].mean() if not hair_data.empty else 0
    hair_avg_days = hair_data['days_diff'].mean() if not hair_data.empty else 0
    hair_mode_weeks = hair_data['weeks_diff'].mode() if not hair_data.empty else pd.Series([0])
    hair_mode_days = hair_data['days_diff'].mode() if not hair_data.empty else pd.Series([0])
    
    # Función para formatear la moda
    def format_mode(mode_series):
        if len(mode_series) == 0:
            return "No data"
        elif len(mode_series) == 1:
            return f"{mode_series.iloc[0]:.2f}"
        else:
            return ", ".join([f"{val:.2f}" for val in mode_series.head(3)])
    
    # ANÁLISIS DE MÚLTIPLES PAYMENT ERRORS
    def count_subs_with_more_than_n(payment_errors_series, max_threshold=10):
        results = {}
        for n in range(2, max_threshold + 1):
            count = (payment_errors_series >= n).sum()
            results[n] = count
        return results
    
    payment_error_counts = count_subs_with_more_than_n(result_df['payment_errors'])
    
    # Crear DataFrames para las hojas de Excel
    
    # Hoja 1: Detalles de suscripciones - ELIMINADA columna days_since_last_frequency_change
    details_df = result_df[[
        'subscription_id', 
        'legacy_category',
        'days_diff', 
        'weeks_diff',
        'frequency_weeks',
        'snooze',
        'payment_errors',
        'has_valid_frequency_change',
        'frequency_changes_json'
    ]].rename(columns={
        'days_diff': 'days_between_orders',
        'weeks_diff': 'weeks_between_orders',
        'has_valid_frequency_change': 'frequency_change_in_period',
        'frequency_changes_json': 'frequency_changes_history'
    })
    
    # Hoja 2: Comparación de frecuencias
    comparison_df = pd.DataFrame({
        'Metric': [
            'Configured Frequency (avg)',
            'Configured Frequency (mode)',
            'Actual Frequency (avg)',
            'Actual Frequency (mode)',
            'BEARD Frequency (avg)',
            'BEARD Frequency (mode)',
            'HAIR Frequency (avg)',
            'HAIR Frequency (mode)',
            'Average Frequency Change (weeks)'
        ],
        'Weeks': [
            freq_avg_weeks,
            format_mode(freq_mode_weeks),
            general_avg_weeks,
            format_mode(general_mode_weeks),
            beard_avg_weeks,
            format_mode(beard_mode_weeks),
            hair_avg_weeks,
            format_mode(hair_mode_weeks),
            f"{avg_frequency_change:.2f}"
        ],
        'Days': [
            freq_avg_weeks * 7,
            format_mode(freq_mode_weeks * 7) if not freq_mode_weeks.empty else "No data",
            general_avg_days,
            format_mode(general_mode_days),
            beard_avg_days,
            format_mode(beard_mode_days),
            hair_avg_days,
            format_mode(hair_mode_days),
            f"{avg_frequency_change * 7:.2f}" if avg_frequency_change else "No data"
        ]
    })
    
    # Hoja 3: Resumen de snooze, payment errors y frequency changes
    summary_table = pd.DataFrame({
        'Metric': [
            'Total Subscriptions',
            'Subscriptions with Snooze',
            'Subscriptions with Payment Errors',
            'Subscriptions with Valid Frequency Changes',
            'Snooze Percentage',
            'Payment Errors Percentage',
            'Frequency Changes Percentage',
            'Average Frequency Change (weeks)'
        ],
        'Value': [
            total_subscriptions,
            total_with_snooze,
            total_with_payment_errors,
            total_with_valid_frequency_changes,
            f"{snooze_percentage:.2f}%",
            f"{payment_errors_percentage:.2f}%",
            f"{frequency_changes_percentage:.2f}%",
            f"{avg_frequency_change:.2f} weeks"
        ]
    })
    
    # Hoja 4: Múltiples payment errors
    multiples_table = pd.DataFrame({
        'Threshold': list(payment_error_counts.keys()),
        'Subscriptions with >N Payment Errors': list(payment_error_counts.values())
    })
    
    # Crear archivo Excel
    with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
        # Hoja de detalles
        details_df.to_excel(writer, sheet_name='Subscription Details', index=False)
        
        # Hoja de comparación de frecuencias
        comparison_df.to_excel(writer, sheet_name='Frequency Comparison', index=False)
        
        # Hoja de resumen
        summary_table.to_excel(writer, sheet_name='Snooze & Payment Summary', index=False)
        
        # Hoja de múltiples ocurrencias
        multiples_table.to_excel(writer, sheet_name='Multiple Payment Errors', index=False)
        
        # Ajustar formatos automáticamente al contenido
        workbook = writer.book
        
        # Función para ajustar ancho de columnas
        def auto_adjust_columns(worksheet, df):
            for idx, col in enumerate(df.columns):
                series = df[col].astype(str)
                max_len = max(
                    series.str.len().max() if len(series) > 0 else 0,
                    len(str(col))
                ) + 2
                worksheet.set_column(idx, idx, max_len)
        
        # Aplicar a todas las hojas
        auto_adjust_columns(writer.sheets['Subscription Details'], details_df)
        auto_adjust_columns(writer.sheets['Frequency Comparison'], comparison_df)
        auto_adjust_columns(writer.sheets['Snooze & Payment Summary'], summary_table)
        auto_adjust_columns(writer.sheets['Multiple Payment Errors'], multiples_table)
    
    print(f"Reporte generado: {fileName}")
    print(f"Suscripciones analizadas: {len(result_df)}")
    print(f"Frecuencia promedio real: {general_avg_weeks:.2f} semanas")
    print(f"Frecuencia moda real: {format_mode(general_mode_weeks)} semanas")
    print(f"Suscripciones con snooze: {total_with_snooze} ({snooze_percentage:.2f}%)")
    print(f"Suscripciones con payment errors: {total_with_payment_errors} ({payment_errors_percentage:.2f}%)")
    print(f"Suscripciones con cambios de frecuencia válidos en el período: {total_with_valid_frequency_changes} ({frequency_changes_percentage:.2f}%)")
    print(f"Cambio promedio de frecuencia: {avg_frequency_change:.2f} semanas")

def realRenewalFrequency(start_date, end_date, folder_name, fc, name):
    # Consulta SQL principal actualizada
    query_main = f"""
    WITH last_orders AS (
        SELECT
            fo.subscription_id,
            fo.created_at,
            ROW_NUMBER() OVER (PARTITION BY fo.subscription_id ORDER BY fo.created_at DESC) AS rn
        FROM bi.fact_orders fo
        JOIN prod_sales_and_subscriptions.sales_orders so ON fo.id = so.id
        WHERE fo.status != 'CANCELLED'
        {fc}
        AND fo.created_at > '{start_date}'
        AND fo.created_at < '{end_date}'
    )
    SELECT
        su.id as subscription_id,    
        sv.legacy_category,
        sv.delivery_frequency,
        su.additionalFields->>"$.snooze" AS snooze,

        -- Últimas dos fechas de órdenes
        MAX(CASE WHEN lo.rn = 1 THEN lo.created_at END) AS last_order_date,
        MAX(CASE WHEN lo.rn = 2 THEN lo.created_at END) AS second_last_order_date,

        -- Diferencia en días entre ellas
        DATEDIFF(
            MAX(CASE WHEN lo.rn = 1 THEN lo.created_at END),
            MAX(CASE WHEN lo.rn = 2 THEN lo.created_at END)
        ) AS days_diff

    FROM prod_sales_and_subscriptions.subscriptions su
    JOIN last_orders lo 
        ON su.id = lo.subscription_id AND lo.rn <= 2
    JOIN prod_sales_and_subscriptions.subscriptions_view sv 
        ON su.id = sv.subscription_id
    GROUP BY 
        su.id, 
        sv.legacy_category, 
        sv.delivery_frequency, 
        su.additionalFields->>"$.snooze"
    """

    saveFile(folder_name, f'renewal_frequency_paymetErrors_{name}.xlsx', query_main)

def saveFile(folder_name, file_name, query_main):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    # Construir la ruta completa del archivo
    full_path = os.path.join(folder_name, file_name)
    renewalFrequency(query_main, full_path)
    #upload_to_drive(full_path, folder_id="1F1VZxlp5IxkQEo4WD0Bt8VEJZ28OhGut")
    #upload_to_dropbox(full_path, dropbox_path=f"/MyReports/{folder_name}/{file_name}")

noFullControl = f"""
    AND (
        so.additionalFields ->> "$.sms_renewal" = "false" 
        OR so.additionalFields ->> "$.sms_renewal" IS NULL
    ) 
"""

fullControl = f"""
    AND so.additionalFields ->> "$.sms_renewal" = "true"      
"""

realRenewalFrequency('2023-01-01', '2024-01-01', 'rw', noFullControl, 'no_fc_2023')
realRenewalFrequency('2023-01-01', '2024-01-01', 'rw', fullControl, 'fc_2023')
realRenewalFrequency('2023-01-01', '2024-01-01', 'rw', "", "2023")

realRenewalFrequency('2024-01-01', '2025-01-01', 'rw', noFullControl, 'no_fc_2024')
realRenewalFrequency('2024-01-01', '2025-01-01', 'rw', fullControl, 'fc_2024')
realRenewalFrequency('2024-01-01', '2025-01-01', 'rw', "", "2024")

realRenewalFrequency('2025-01-01', '2025-11-01', 'rw', noFullControl, 'no_fc_2025')
realRenewalFrequency('2025-01-01', '2025-11-01', 'rw', fullControl, 'fc_2025')
realRenewalFrequency('2025-01-01', '2025-11-01', 'rw', "", "2025")