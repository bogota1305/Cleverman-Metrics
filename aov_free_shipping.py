import os
import pandas as pd
import numpy as np # Necesitamos numpy para np.inf
from modules.database_queries import execute_query
from uploadCloud import upload_to_drive

def clean_amount(value):
    try:
        return float(value)
    except:
        if isinstance(value, str):
            clean_val = value.split('.')[0]
            try:
                return float(clean_val)
            except:
                return None
        return None

def get_total_orders(start_date, end_date, is_subscribed):
    query = f"""
    SELECT COUNT(*) AS total_orders
    FROM 
        prod_sales_and_subscriptions.intents
    WHERE 
        createdAt < '{end_date}'
        AND createdAt > '{start_date}'
        AND (content ->> '$.additionalFields.subscribed' = '{str(is_subscribed).lower()}')
        AND (content ->> '$.status' != 'CANCELLED')
        AND orderNumber IS NOT NULL;
    """
    result = execute_query(query)

    # Manejar diferentes tipos de retorno de execute_query
    if isinstance(result, pd.DataFrame):
        return result.iloc[0, 0] if not result.empty else 0
    elif isinstance(result, (list, tuple)) and len(result) > 0:
        return result[0][0]
    else:
        return 0

# --- Modificación: Pasar is_subscribed a _process_sheet ---
def generate_order_report(query, filename, start_date, end_date, is_subscribed):
    # Ejecutar la consulta SQL principal
    data = execute_query(query)
    df = pd.DataFrame(data, columns=['orderNumber', 'createdAt', 'total_amount', 'units', 'content']) # Añadimos 'content' temporalmente si es necesario para futuros análisis, aunque no se usa en este procesamiento directo
    
    # Obtener total de órdenes para el período y tipo
    total_orders = get_total_orders(start_date, end_date, is_subscribed)
    
    # Convertir fechas
    df['createdAt'] = pd.to_datetime(df['createdAt'])
    
    # Limpiar y convertir valores numéricos
    df['total_amount'] = df['total_amount'].apply(clean_amount)
    df = df.dropna(subset=['total_amount'])
    
    df['units'] = pd.to_numeric(df['units'], errors='coerce')
    df = df.dropna(subset=['units'])
    
    # Separar en órdenes e intentos
    orders_df = df[df['orderNumber'].notnull()].copy()
    attempts_df = df[df['orderNumber'].isnull()].copy()
    
    # Calcular total de intentos (todos los registros devueltos por la query)
    total_attempts = len(df)
    
    # Crear el archivo Excel
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # Ahora pasamos is_subscribed a _process_sheet
        _process_sheet(orders_df, writer, 'Órdenes', total_orders, start_date, end_date, total_attempts, is_subscribed)
        _process_sheet(attempts_df, writer, 'Intentos', total_orders, start_date, end_date, total_attempts, is_subscribed)
    
    print(f"Reporte generado exitosamente: {filename}")

# --- Modificación principal: Función _process_sheet ---
# Añadimos el parámetro is_subscribed
def _process_sheet(df, writer, sheet_name, total_orders=None, start_date=None, end_date=None, total_attempts=None, is_subscribed=False):
    if len(df) == 0:
        # Si el DataFrame está vacío, escribir una hoja vacía con los encabezados principales
        empty_summary = pd.DataFrame(columns=['Métrica', 'Valor'])
        empty_price_ranges = pd.DataFrame(columns=['Rango de Precio', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones'])
        empty_items = pd.DataFrame(columns=['Items', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones'])

        with pd.ExcelWriter(writer, engine='xlsxwriter', mode='a') as writer:
            empty_summary.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)
            empty_price_ranges.to_excel(writer, sheet_name=sheet_name, startrow=len(empty_summary) + 2, index=False)
            empty_items.to_excel(writer, sheet_name=sheet_name, startrow=len(empty_summary) + 2 + len(empty_price_ranges) + 2, index=False)

        return

    # Total de transacciones en este DataFrame (será la base para los porcentajes en tablas secundarias)
    total_transactions_in_df = len(df)

    # Calcular métricas básicas
    avg_amount = df['total_amount'].mean() if not df['total_amount'].empty else 0
    avg_units = df['units'].mean() if not df['units'].empty else 0

    # Calcular el rango de fechas y total de días del dataset específico
    if not df['createdAt'].empty and pd.notna(df['createdAt'].min()) and pd.notna(df['createdAt'].max()):
        min_date = df['createdAt'].min().date()
        max_date = df['createdAt'].max().date()
        filtered_days = (max_date - min_date).days + 1
    else:
        min_date, max_date = None, None
        filtered_days = 0

    # Calcular promedio de transacciones específicas por día
    specific_transactions_per_day = total_transactions_in_df / filtered_days if filtered_days > 0 else 0

    # Inicializar variables para métricas adicionales
    total_orders_per_day = None
    percentage_of_total = None # Porcentaje de este DF sobre el total diario de órdenes
    conversion_rate = None # Tasa de conversión de la query a órdenes (solo en Intentos)

    # Solo para la hoja de Órdenes
    if sheet_name == 'Órdenes' and total_orders is not None and start_date and end_date:
        start_dt = pd.to_datetime(start_date).date()
        end_dt = pd.to_datetime(end_date).date()
        full_period_days = (end_dt - start_dt).days + 1
        total_orders_per_day = total_orders / full_period_days if full_period_days > 0 else 0
        # Calculamos este porcentaje para la fila del summary_df
        percentage_of_total = (specific_transactions_per_day / total_orders_per_day) if total_orders_per_day > 0 else 0

    # Solo para la hoja de Intentos
    if sheet_name == 'Intentos' and total_attempts is not None:
        # total_attempts es el total de filas de la query principal
        # len(df) en la hoja Intentos es el número de INTENTOS FALLIDOS de la query
        # Las órdenes exitosas de la query principal son total_attempts - (Intentos fallidos + Órdenes canceladas que pasaron el filtro + otros casos raros)
        # Una aproximación es considerar que las órdenes exitosas son las que NO están en este df de intentos fallidos
        successful_orders = total_attempts - len(df) # Esta es una estimación simplificada
        conversion_rate = successful_orders / total_attempts if total_attempts > 0 else 0 # Tasa de conversión de la query a una estimación de órdenes exitosas

    # --- Resumen por rangos de precio ---
    price_ranges_summary = pd.DataFrame(columns=['Rango de Precio', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones'])
    if not df['total_amount'].empty and df['total_amount'].dropna().any(): # Solo crear rangos si hay montos válidos
        # Definir rangos y etiquetas según si está suscrito
        if is_subscribed:
            bins = [0, 15, 30, 45, 60, 75, 90, np.inf]
            labels = ['0 - 15', '15 - 30', '30 - 45', '45 - 60', '60 - 75', '75 - 90', '> 90']
        else:
            bins = [0, 20, 35, 50, 65, 80, 95, np.inf]
            labels = ['0 - 20', '20 - 35', '35 - 50', '50 - 65', '65 - 80', '80 - 95', '> 95']

        # Crear la columna de rango de precio usando pd.cut
        df['price_range'] = pd.cut(df['total_amount'], bins=bins, labels=labels, right=True, include_lowest=True)

        # Agrupar por el nuevo rango de precio y calcular métricas
        # Usamos size para contar las filas en cada grupo
        price_ranges_summary = df.groupby('price_range', observed=False).agg(
            total_transactions=('orderNumber', 'size'), # size cuenta filas
            mean_amount=('total_amount', 'mean')
        ).reset_index()

        price_ranges_summary.columns = ['Rango de Precio', 'Total Transacciones', 'Precio Promedio']

        # Calcular 'Transacciones por día' para cada rango
        price_ranges_summary['Transacciones por día'] = price_ranges_summary['Total Transacciones'] / filtered_days if filtered_days > 0 else 0

        # Calcular 'Porcentaje de Transacciones' para cada rango
        price_ranges_summary['Porcentaje de Transacciones'] = price_ranges_summary['Total Transacciones'] / total_transactions_in_df if total_transactions_in_df > 0 else 0

        # Reordenar columnas
        price_ranges_summary = price_ranges_summary[['Rango de Precio', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones']]


    # --- Resumen por cantidad de items ---
    items_summary = pd.DataFrame(columns=['Items', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones'])
    if not df['units'].empty and df['units'].dropna().any(): # Solo crear resumen de items si hay unidades válidas
        items_summary = df.groupby('units', observed=False).agg(
            total_transactions=('units', 'size'), # size cuenta filas
            mean_amount=('total_amount', 'mean')
        ).reset_index()

        items_summary.columns = ['Items', 'Total Transacciones', 'Precio Promedio']

        # Calcular 'Transacciones por día' para cada cantidad de ítems
        items_summary['Transacciones por día'] = items_summary['Total Transacciones'] / filtered_days if filtered_days > 0 else 0

        # Nuevo cálculo: Porcentaje de transacciones para cada cantidad de ítems
        items_summary['Porcentaje de Transacciones'] = items_summary['Total Transacciones'] / total_transactions_in_df if total_transactions_in_df > 0 else 0

        # Reordenar columnas
        items_summary = items_summary[['Items', 'Total Transacciones', 'Precio Promedio', 'Transacciones por día', 'Porcentaje de Transacciones']]


    # --- Crear el resumen con las métricas básicas y adicionales ---
    # Se trae de vuelta la fila de Porcentaje de Órdenes sobre total diario para la hoja 'Órdenes'
    summary_data = {
        'Métrica': ['Precio Promedio', 'Items Promedio', f'Transacciones por día (promedio)'],
        'Valor': [avg_amount, avg_units, specific_transactions_per_day]
    }

    # Agregar métricas adicionales para Órdenes
    if sheet_name == 'Órdenes' and total_orders_per_day is not None:
        summary_data['Métrica'].extend([
            'Órdenes totales por día (promedio)',
            f'Porcentaje de {sheet_name} sobre total diario' # Se añade de nuevo
        ])
        summary_data['Valor'].extend([
            total_orders_per_day,
            percentage_of_total # Este valor es necesario para la fila
        ])

    # Agregar métrica de conversión para Intentos
    if sheet_name == 'Intentos' and conversion_rate is not None:
         summary_data['Métrica'].append(f'Tasa de conversión de la query a Órdenes')
         summary_data['Valor'].append(conversion_rate)

    summary_df = pd.DataFrame(summary_data)

    # Escribir los DataFrames en el Excel
    current_startrow = 0
    summary_df.to_excel(writer, sheet_name=sheet_name, startrow=current_startrow, index=False)

    # Calcular la fila de inicio para las tablas secundarias después de escribir el summary_df
    price_range_start_row_in_sheet = len(summary_df) + current_startrow + 2
    price_ranges_summary.to_excel(writer, sheet_name=sheet_name,
                                  startrow=price_range_start_row_in_sheet, index=False)

    items_start_row_in_sheet = price_range_start_row_in_sheet + len(price_ranges_summary) + 2
    items_summary.to_excel(writer, sheet_name=sheet_name,
                           startrow=items_start_row_in_sheet, index=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Formatos numéricos y de porcentaje
    num_format = workbook.add_format({'num_format': '#,##0.00'}) # Formato general con 2 decimales
    percent_format = workbook.add_format({'num_format': '0.00%'}) # Formato de porcentaje
    int_format = workbook.add_format({'num_format': '0'}) # Formato de entero

    # --- Ajuste de anchos de columnas y formatos ---
    # Columna A: Métrica / Rango de Precio / Items
    list_col_A_lengths = [len(str(x)) for x in summary_df['Métrica']]
    if not price_ranges_summary.empty:
         list_col_A_lengths.extend([len(str(x)) for x in price_ranges_summary['Rango de Precio'].astype(str)])
    if not items_summary.empty:
         list_col_A_lengths.extend([len(str(x)) for x in items_summary['Items'].astype(str)])
    max_len_col_A = max(list_col_A_lengths) + 2 if list_col_A_lengths else 15
    worksheet.set_column(0, 0, max_len_col_A)

    # Columna B: Valor (Summary) / Total Transacciones (Rango) / Total Transacciones (Items)
    # Incluir longitudes de los encabezados 'Valor' y 'Total Transacciones'
    list_col_B_lengths = [len('Valor'), len('Total Transacciones')]
    if not summary_df.empty:
         list_col_B_lengths.extend([len(str(f"{x:.2f}")) for x in summary_df['Valor'] if pd.notnull(x)])
    if not price_ranges_summary.empty:
         list_col_B_lengths.extend([len(str(int(x))) for x in price_ranges_summary['Total Transacciones']])
    if not items_summary.empty:
         list_col_B_lengths.extend([len(str(int(x))) for x in items_summary['Total Transacciones']])
    max_len_col_B = max(list_col_B_lengths) + 2 if list_col_B_lengths else 15
    worksheet.set_column(1, 1, max_len_col_B, num_format) # Aplicaremos los formatos específicos celda a celda/rango a rango

    # Columna C: Precio Promedio (Rango) / Precio Promedio (Items)
    list_col_C_lengths = [len('Precio Promedio')]
    if not price_ranges_summary.empty:
        list_col_C_lengths.extend([len(str(f"{x:.2f}")) for x in price_ranges_summary['Precio Promedio'] if pd.notnull(x)])
    if not items_summary.empty:
        list_col_C_lengths.extend([len(str(f"{x:.2f}")) for x in items_summary['Precio Promedio'] if pd.notnull(x)])
    max_len_col_C = max(list_col_C_lengths) + 2 if list_col_C_lengths else 15
    worksheet.set_column(2, 2, max_len_col_C, num_format)

    # Columna D: Transacciones por día (Rango) / Transacciones por día (Items)
    list_col_D_lengths = [len('Transacciones por día')]
    if not price_ranges_summary.empty:
        list_col_D_lengths.extend([len(str(f"{x:.2f}")) for x in price_ranges_summary['Transacciones por día'] if pd.notnull(x)])
    if not items_summary.empty:
        list_col_D_lengths.extend([len(str(f"{x:.2f}")) for x in items_summary['Transacciones por día'] if pd.notnull(x)])
    max_len_col_D = max(list_col_D_lengths) + 2 if list_col_D_lengths else 15
    worksheet.set_column(3, 3, max_len_col_D, num_format)

    # Columna E: Porcentaje de Transacciones (Rango) / Porcentaje de Transacciones (Items)
    list_col_E_lengths = [len('Porcentaje de Transacciones')]
    if not price_ranges_summary.empty:
         list_col_E_lengths.extend([len("{:.2f}%".format(x * 100).replace('%', '')) for x in price_ranges_summary['Porcentaje de Transacciones'] if pd.notnull(x)])
    if not items_summary.empty:
         list_col_E_lengths.extend([len("{:.2f}%".format(x * 100).replace('%', '')) for x in items_summary['Porcentaje de Transacciones'] if pd.notnull(x)])
    max_len_col_E = max(list_col_E_lengths) + 2 if list_col_E_lengths else 15
    worksheet.set_column(4, 4, max_len_col_E, percent_format)


    # --- Aplicar formatos especiales a filas específicas del summary básico ---
    # Aplicar formato numérico a Órdenes totales por día (si existe)
    selected_index_total_daily = summary_df[summary_df['Métrica'] == 'Órdenes totales por día (promedio)'].index
    if not selected_index_total_daily.empty:
        row_idx_total_daily_orders = selected_index_total_daily[0]
        # Aplicar formato en la fila row_idx + 1 (basado en corrección anterior)
        worksheet.write(row_idx_total_daily_orders + 1, 1, total_orders_per_day, num_format)

    # Aplicar formato porcentual a la métrica de porcentaje sobre total diario (si existe en sheet 'Órdenes')
    if sheet_name == 'Órdenes':
         selected_index_percentage_total_daily = summary_df[summary_df['Métrica'] == f'Porcentaje de {sheet_name} sobre total diario'].index
         if not selected_index_percentage_total_daily.empty:
             row_idx_percentage_total_daily = selected_index_percentage_total_daily[0]
             # Aplicar formato en la fila row_idx + 1 (basado en corrección anterior)
             worksheet.write(row_idx_percentage_total_daily + 1, 1, percentage_of_total, percent_format)


    # Aplicar formato porcentual a la métrica de conversión en Intentos (si existe)
    selected_index_conversion = summary_df[summary_df['Métrica'] == f'Tasa de conversión de la query a Órdenes'].index
    if not selected_index_conversion.empty:
         row_idx_conversion = selected_index_conversion[0]
         # Aplicar formato en la fila row_idx + 1 (corrección solicitada)
         worksheet.write(row_idx_conversion + 1, 1, conversion_rate, percent_format)


    # --- Aplicar formato entero a la columna 'Total Transacciones' en las tablas secundarias ---
    col_idx_total_transactions = 1 # La columna 'Total Transacciones' es la segunda columna (índice 1)

    # Rango de filas para price_ranges_summary (desde la fila de datos hasta el final)
    # La fila de inicio de datos es la fila del encabezado + 1
    price_range_data_start_row = price_range_start_row_in_sheet + 1
    if not price_ranges_summary.empty:
        for i, row_data in enumerate(price_ranges_summary['Total Transacciones']):
            # Escribimos el valor de nuevo con el formato entero
            worksheet.write(price_range_data_start_row + i, col_idx_total_transactions, row_data, int_format)

    # Rango de filas para items_summary (desde la fila de datos hasta el final)
    items_data_start_row = items_start_row_in_sheet + 1
    if not items_summary.empty:
        for i, row_data in enumerate(items_summary['Total Transacciones']):
             # Escribimos el valor de nuevo con el formato entero
             worksheet.write(items_data_start_row + i, col_idx_total_transactions, row_data, int_format)

# --- Funciones existentes (sin cambios) ---
def getQuerry(start_date, end_date, is_subscribed):
    # La query original ya filtra por el monto mínimo, lo cual es perfecto para los rangos
    return f"""
SELECT 
    orderNumber, 
    createdAt,
    (
        SELECT SUM(JSON_EXTRACT(item.value, '$.unitPrice') * JSON_EXTRACT(item.value, '$.quantity'))
        FROM JSON_TABLE(content, '$.items[*]' COLUMNS(
            value JSON PATH '$'
        )) AS item
        WHERE JSON_EXTRACT(item.value, '$.unitPrice') > 0
    ) - 
    (
        SELECT IFNULL(SUM(JSON_EXTRACT(discount.value, '$.value')), 0)
        FROM JSON_TABLE(content, '$.discounts[*]' COLUMNS(
            value JSON PATH '$'
        )) AS discount
        WHERE JSON_EXTRACT(discount.value, '$.additionalFields.type') != 'shipping'
        OR JSON_EXTRACT(discount.value, '$.additionalFields.type') IS NULL
    ) AS total_amount,
    (
        SELECT SUM(JSON_EXTRACT(item.value, '$.quantity'))
        FROM JSON_TABLE(content, '$.items[*]' COLUMNS(
            value JSON PATH '$'
        )) AS item
        WHERE JSON_EXTRACT(item.value, '$.unitPrice') > 0
    ) AS units,
    content -- Incluimos content por si acaso, aunque no se procese directamente en _process_sheet
FROM 
    prod_sales_and_subscriptions.intents
WHERE 
    createdAt > '{start_date}'
    AND createdAt < '{end_date}'
    AND (
        SELECT SUM(JSON_EXTRACT(item.value, '$.unitPrice') * JSON_EXTRACT(item.value, '$.quantity'))
        FROM JSON_TABLE(content, '$.items[*]' COLUMNS(
            value JSON PATH '$'
        )) AS item
        WHERE JSON_EXTRACT(item.value, '$.unitPrice') > 0
    ) - 
    (
        SELECT IFNULL(SUM(JSON_EXTRACT(discount.value, '$.value')), 0)
        FROM JSON_TABLE(content, '$.discounts[*]' COLUMNS(
            value JSON PATH '$'
        )) AS discount
        WHERE JSON_EXTRACT(discount.value, '$.additionalFields.type') != 'shipping'
        OR JSON_EXTRACT(discount.value, '$.additionalFields.type') IS NULL
    ) >= '{35 if not is_subscribed else 15}'    
    AND (content ->> '$.additionalFields.subscribed' = '{str(is_subscribed).lower()}')
    AND (content ->> '$.status' != 'CANCELLED');
"""

def getQuerryB(start_date, end_date, is_subscribed):
    # La query original ya filtra por el monto mínimo, lo cual es perfecto para los rangos
    return f"""
SELECT 
    orderNumber, 
    createdAt,
    (
        SELECT SUM(JSON_EXTRACT(item.value, '$.unitPrice') * JSON_EXTRACT(item.value, '$.quantity'))
        FROM JSON_TABLE(content, '$.items[*]' COLUMNS(
            value JSON PATH '$'
        )) AS item
        WHERE JSON_EXTRACT(item.value, '$.unitPrice') > 0
    ) - 
    (
        SELECT IFNULL(SUM(JSON_EXTRACT(discount.value, '$.value')), 0)
        FROM JSON_TABLE(content, '$.discounts[*]' COLUMNS(
            value JSON PATH '$'
        )) AS discount
        WHERE JSON_EXTRACT(discount.value, '$.additionalFields.type') != 'shipping'
        OR JSON_EXTRACT(discount.value, '$.additionalFields.type') IS NULL
    ) AS total_amount,
    (
        SELECT SUM(JSON_EXTRACT(item.value, '$.quantity'))
        FROM JSON_TABLE(content, '$.items[*]' COLUMNS(
            value JSON PATH '$'
        )) AS item
        WHERE JSON_EXTRACT(item.value, '$.unitPrice') > 0
    ) AS units,
    content -- Incluimos content por si acaso, aunque no se procese directamente en _process_sheet
FROM 
    prod_sales_and_subscriptions.intents
WHERE 
    createdAt > '{start_date}'
    AND createdAt < '{end_date}' 
    AND (content ->> '$.additionalFields.subscribed' = '{str(is_subscribed).lower()}')
    AND (content ->> '$.status' != 'CANCELLED');
"""

# --- Modificación: Pasar is_subscribed a generate_order_report ---
def saveFile(file_name, start_date, end_date, is_subscribed):
    folder_name = "FreeShipping"

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    full_path = os.path.join(folder_name, f"{file_name}_{start_date}_{end_date}.xlsx")
    # Pasamos is_subscribed a getQuerry y generate_order_report
    query = getQuerry(start_date, end_date, is_subscribed)
    # query = getQuerryB(start_date, end_date, is_subscribed)
    generate_order_report(query, full_path, start_date, end_date, is_subscribed) # Pasamos is_subscribed aquí también
    upload_to_drive(full_path, folder_id="1F1VZxlp5IxkQEo4WD0Bt8VEJZ28OhGut")

# Generar reportes con los parámetros adicionales (sin cambios en esta sección)
saveFile('oto_free_shipping.xlsx', '2025-02-20', '2025-07-21', False)
saveFile('sub_free_shipping.xlsx', '2025-02-20', '2025-07-21', True)
# saveFile('oto_no_free_shipping.xlsx', '2024-12-20', '2025-02-20', False)
# saveFile('sub_no_free_shipping.xlsx', '2024-12-20', '2025-02-20', True)

# saveFile('oto_free_shipping_no_rule', '2025-02-20', '2025-07-14', False)
# saveFile('sub_free_shipping_no_rule', '2025-02-20', '2025-07-14', True)
# saveFile('oto_no_free_shipping_no_rule', '2024-12-20', '2025-07-14', False)
# saveFile('sub_no_free_shipping_no_rule', '2024-12-20', '2025-07-14', True)