import pandas as pd
import mysql.connector
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill
from tkinter import Tk, Label, Button, Entry
from tkinter import messagebox
from tkcalendar import Calendar
from modules.colors import lighten_color
from modules.database_queries import execute_query
from modules.date_selector import open_date_selector
from modules.excel_creator import save_dataframe_to_excel, save_dataframe_to_excel_orders
from report import anotar_datos_excel

def consulta(start_date, end_date):

    query_orders = f"""
    SELECT 
        fo.*,
        CASE
            WHEN MAX(CASE WHEN fsoi.itemId = 'IT00000000000000000000000000000107' THEN 1 ELSE 0 END) = 1 
            THEN fo.units - SUM(CASE WHEN fsoi.itemId = 'IT00000000000000000000000000000107' THEN fsoi.quantity ELSE 0 END)
            ELSE fo.units
        END AS validItems
    FROM 
        bi.fact_orders fo
    LEFT JOIN 
        bi.fact_sales_order_items fsoi ON fsoi.salesOrderId = fo.id
    WHERE
        fo.created_at >= '{start_date} 00:00:00' 
        AND fo.created_at < '{end_date} 00:00:00' 
        AND fo.status NOT IN ('CANCELLED', 'PAYMENT_ERROR') 
    GROUP BY 
        fo.id, fo.order_number, fo.units;  
    """

    sO = execute_query(query_orders)

    sO_sin_recurrentes = sO[sO['recurrent'] == 0]
    
    # Crear grupos por tipo de usuario
    usuarios_nuevos = sO_sin_recurrentes[sO_sin_recurrentes['is_first_order'] == 1]
    usuarios_antiguos = sO_sin_recurrentes[sO_sin_recurrentes['is_first_order'] == 0]

    usuarios_nuevos_sub = usuarios_nuevos[usuarios_nuevos['order_plan'] == 'SUBSCRIPTION']
    usuarios_nuevos_oto = usuarios_nuevos[usuarios_nuevos['order_plan'] == 'OTO']
    usuarios_nuevos_mix = usuarios_nuevos[usuarios_nuevos['order_plan'] == 'MIXED']
    
    usuarios_antiguos_sub = usuarios_antiguos[usuarios_antiguos['order_plan'] == 'SUBSCRIPTION']
    usuarios_antiguos_oto = usuarios_antiguos[usuarios_antiguos['order_plan'] == 'OTO']
    usuarios_antiguos_mix = usuarios_antiguos[usuarios_antiguos['order_plan'] == 'MIXED']

    sO_con_recurrentes = sO[sO['recurrent'] == 1]

    return usuarios_nuevos_sub, usuarios_nuevos_oto, usuarios_nuevos_mix, usuarios_antiguos_sub, usuarios_antiguos_oto, usuarios_antiguos_mix, sO_con_recurrentes, usuarios_antiguos, usuarios_nuevos

def process_data(metrica, start_date, end_date):
    # Convertir fechas
    metrica['created_at'] = pd.to_datetime(metrica['created_at'])
    metrica['date'] = metrica['created_at'].dt.date

    # Crear un rango completo de fechas
    all_dates = pd.date_range(start=start_date, end=end_date).date
    all_dates_df = pd.DataFrame({'date': all_dates})

    # Cálculo de valores diarios
    daily_summary = metrica.groupby('date').agg(
        total_revenue=('total', 'sum'),
        total_sales=('order_number', 'count'),
        total_items=('validItems', 'sum'),
    ).reset_index()

    # Fusionar con todas las fechas
    daily_summary = pd.merge(all_dates_df, daily_summary, on='date', how='left').fillna(0)

    # Agregar las columnas de promedio
    daily_summary['average_items'] = (
        daily_summary['total_items'] / daily_summary['total_sales'].replace(0, 1)
    ).where(daily_summary['total_sales'] > 0, 0)  # Forzar a 0 si total_sales es 0

    daily_summary['average_value'] = (
        daily_summary['total_revenue'] / daily_summary['total_sales'].replace(0, 1)
    ).where(daily_summary['total_sales'] > 0, 0)  # Forzar a 0 si total_sales es 0

    # Cálculo de totales generales
    suma_total = metrica['total'].sum()
    total_ventas = metrica['total'].count()
    total_items = metrica['validItems'].sum()
    average_items = total_items / total_ventas if total_ventas > 0 else 0
    average_value = suma_total / total_ventas if total_ventas > 0 else 0

    # Resumen total
    totals_row = pd.DataFrame([{
        'date': 'Total',
        'total_revenue': suma_total,
        'total_sales': total_ventas,
        'total_items': total_items,
        'average_items': average_items,
        'average_value': average_value,
    }])

    daily_summary = pd.concat([daily_summary, totals_row], ignore_index=True)

    return daily_summary, average_items, average_value


def documento(resumeData, file_name, folder_name, dropbox_var, drive_var):

    columns_to_plot = ['total_revenue', 'total_sales', 'total_items', 'average_items', 'average_value']
    colors = ['#0000FF', '#008000', '#FF0000', '#6F4E37', '#FF00FF']
    grafico_positions = ['H2', 'H24', 'H46', 'H68', 'H90']

    url = save_dataframe_to_excel_orders(folder_name, file_name, resumeData, 'General', columns_to_plot, colors, grafico_positions, dropbox_var, drive_var)
    return url


def get_orders(start_date, end_date, folder_name, unique_orders_var, dropbox_var, drive_var):

    usuarios_nuevos_sub, usuarios_nuevos_oto, usuarios_nuevos_mix, usuarios_antiguos_sub, usuarios_antiguos_oto, usuarios_antiguos_mix, recurrentes, usuarios_antiguos, usuarios_nuevos = consulta(start_date, end_date)

    #------------------------------------------------------------------------------------------------

    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    new_end_date = (end_date - pd.DateOffset(days=1)).strftime('%Y-%m-%d')
    items = []
    values = []
    urls = []

    if(unique_orders_var[0] == 1):
        data_usuarios_nuevos_sub, average_items, average_value = process_data(usuarios_nuevos_sub, start_date, new_end_date)
        url = documento(data_usuarios_nuevos_sub, f'New users - SUBS', folder_name, dropbox_var, drive_var)
        items.insert(0, average_items)
        values.insert(0, average_value)
        urls.insert(0, url)    

    if(unique_orders_var[1] == 1):
        data_usuarios_nuevos_oto, average_items, average_value = process_data(usuarios_nuevos_oto, start_date, new_end_date)
        url = documento(data_usuarios_nuevos_oto, f'New users - OTO', folder_name, dropbox_var, drive_var)
        items.insert(1, average_items)
        values.insert(1, average_value)
        urls.insert(1, url)

    if(unique_orders_var[2] == 1):
        data_usuarios_nuevos_mix, average_items, average_value = process_data(usuarios_nuevos_mix, start_date, new_end_date)
        url = documento(data_usuarios_nuevos_mix, f'New users - MIX', folder_name, dropbox_var, drive_var)
        items.insert(2, average_items)
        values.insert(2, average_value)
        urls.insert(2, url)

    if(unique_orders_var[3] == 1):
        data_usuarios_nuevos, average_items, average_value = process_data(usuarios_nuevos, start_date, new_end_date)
        url = documento(data_usuarios_nuevos, f'New users - ALL', folder_name, dropbox_var, drive_var)
        items.insert(3, average_items)
        values.insert(3, average_value)
        urls.insert(3, url)

    if(unique_orders_var[4] == 1):
        data_usuarios_antiguos_sub, average_items, average_value = process_data(usuarios_antiguos_sub, start_date, new_end_date)
        url = documento(data_usuarios_antiguos_sub, f'Non recurrent orders of existing users - SUBS', folder_name, dropbox_var, drive_var)
        items.insert(4, average_items)
        values.insert(4, average_value)
        urls.insert(4, url)

    if(unique_orders_var[5] == 1):
        data_usuarios_antiguos_oto, average_items, average_value = process_data(usuarios_antiguos_oto, start_date, new_end_date)
        url = documento(data_usuarios_antiguos_oto, f'Non recurrent orders of existing users - OTO', folder_name, dropbox_var, drive_var)
        items.insert(5, average_items)
        values.insert(5, average_value)
        urls.insert(5, url)

    if(unique_orders_var[6] == 1):
        data_usuarios_antiguos_mix, average_items, average_value = process_data(usuarios_antiguos_mix, start_date, new_end_date)
        url = documento(data_usuarios_antiguos_mix, f'Non recurrent orders of existing users - MIX', folder_name, dropbox_var, drive_var)
        items.insert(6, average_items)
        values.insert(6, average_value)
        urls.insert(6, url)

    if(unique_orders_var[7] == 1):
        data_usuarios_antiguos, average_items, average_value = process_data(usuarios_antiguos, start_date, new_end_date)
        url = documento(data_usuarios_antiguos, f'Non recurrent orders of existing users - ALL', folder_name, dropbox_var, drive_var)
        items.insert(7, average_items)
        values.insert(7, average_value)
        urls.insert(7, url)

    if(unique_orders_var[8] == 1):
        data_recurrentes, average_items, average_value = process_data(recurrentes, start_date, new_end_date)
        url = documento(data_recurrentes, f'Recurrent Orders (ALL)', folder_name, dropbox_var, drive_var)
        items.insert(8, average_items)
        values.insert(8, average_value)
        urls.insert(8, url)

    return values, items, urls




