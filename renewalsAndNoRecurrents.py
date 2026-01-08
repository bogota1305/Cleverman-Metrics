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
from modules.excel_creator import save_dataframe_to_excel

def process_data(start_date, end_date):

    query_orders = f"""
    SELECT *
    FROM bi.fact_orders
    WHERE created_at >= '{start_date} 00:00:00' -- Fecha de inicio
    AND created_at < '{end_date} 00:00:00' -- Fecha actual
    AND status NOT IN ('CANCELLED', 'PAYMENT_ERROR')
    AND is_first_order <> 1;
    """
    sO = execute_query(query_orders)

    sO_sin_recurrentes = sO[sO['recurrent'] == 0]
    sO_con_recurrentes = sO[sO['recurrent'] == 1]

    # Convertir fechas
    sO_sin_recurrentes['created_at'] = pd.to_datetime(sO['created_at'])
    sO_sin_recurrentes['date'] = sO['created_at'].dt.date

    # Convertir fechas
    sO_con_recurrentes['created_at'] = pd.to_datetime(sO['created_at'])
    sO_con_recurrentes['date'] = sO['created_at'].dt.date

    # Cálculo de valores diarios para la primera página
    daily_summary = sO_sin_recurrentes.groupby('date').agg(
        daily_no_recurrent=('recurrent', 'count'),
    ).reset_index()

     # Cálculo de valores diarios para la primera página
    daily_summary_recurrents = sO_con_recurrentes.groupby('date').agg(
        daily_recurrent=('recurrent', 'count'),
    ).reset_index()

     # Combinar las métricas diarias en un solo DataFrame
    daily_summary = pd.merge(
        daily_summary, 
        daily_summary_recurrents, 
        on='date', 
        how='outer'
    ).fillna(0)  # Rellenar valores faltantes con 0

    # Asegurar tipos de datos correctos después de merge
    daily_summary['daily_no_recurrent'] = daily_summary['daily_no_recurrent'].astype(int)
    daily_summary['daily_recurrent'] = daily_summary['daily_recurrent'].astype(int)

    # Cálculo de totales generales
    total_no_recurrent = sO_sin_recurrentes['recurrent'].count()
    total_recurrent = sO_con_recurrentes['recurrent'].count()

    total_no_recurrent_average = daily_summary['daily_no_recurrent'].mean()
    total_recurrent_average = daily_summary['daily_recurrent'].mean()

    # Resumen total
    totals_row = pd.DataFrame([{
        'date': 'Total',
        'daily_no_recurrent': total_no_recurrent,
        'daily_recurrent': total_recurrent
    }])

    total_sales = [total_recurrent_average, total_recurrent, total_no_recurrent_average, total_no_recurrent]

    daily_summary = pd.concat([daily_summary, totals_row], ignore_index=True)

    return daily_summary, total_sales

def get_sales(start_date, end_date, folder_name, dropbox_var, drive_var):
    
    file_name = 'Sales'
    resumeData, total_sales = process_data(start_date, end_date)

    # Crear gráficos para resumen general
    columns_to_plot = ['daily_no_recurrent', 'daily_recurrent']
    colors = ['#0000FF', '#008000']
    grafico_positions = ['H2', 'H24']

    urls = save_dataframe_to_excel(folder_name, file_name, resumeData, 'General', columns_to_plot, colors, grafico_positions, dropbox_var, drive_var)

    return total_sales, urls