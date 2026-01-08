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

def subs(start_date, end_date):

    query_minisubs = f"""
    SELECT * FROM bi.fact_subscriptions
	WHERE created_at >= '{start_date} 00:00:00' -- Fecha de inicio
    AND created_at < '{end_date} 00:00:00' -- Fecha actual
    AND status != 'CANCELLED'
    AND plan_id IN ('SP00000000000000000000000000000012', 'SP00000000000000000000000000000013', 'SP00000000000000000000000000000014', 'SP00000000000000000000000000000015', 'SP00000000000000000000000000000016', 'SP00000000000000000000000000000017', 'SP00000000000000000000000000000018', 'SP00000000000000000000000000000019', 'SP00000000000000000000000000000020', 'SP00000000000000000000000000000021', 'SP00000000000000000000000000000022', 'SP00000000000000000000000000000023');
    """

    query_subs = f"""
    SELECT * FROM bi.fact_subscriptions
	WHERE created_at >= '{start_date} 00:00:00' -- Fecha de inicio
    AND created_at < '{end_date} 00:00:00' -- Fecha actual
    AND status != 'CANCELLED'
    AND plan_id IN ('SP00000000000000000000000000000002', 'SP00000000000000000000000000000003', 'SP00000000000000000000000000000004', 'SP00000000000000000000000000000005', 'SP00000000000000000000000000000008', 'SP00000000000000000000000000000009', 'SP00000000000000000000000000000010', 'SP00000000000000000000000000000011');
    """

    query_oto = f"""
    SELECT *
    FROM bi.fact_sales_order_items so
    JOIN bi.fact_orders fo ON so.salesOrderId = fo.id
    WHERE fo.created_at >= '2025-06-01 00:00:00' -- Fecha de inicio
    AND fo.created_at < '2025-07-01 00:00:00' -- Fecha actual
    AND fo.status != 'CANCELLED'
    AND fo.order_plan != 'SUBSCRIPTION'
    AND so.itemId IN ('IT00000000000000000000001004170001', 'IT00000000000000000000001004170002', 'IT00000000000000000000001004170003', 'IT00000000000000000000001004170004', 'IT00000000000000000000001004170005', 'IT00000000000000000000001004170006', 'IT00000000000000000000001004170007', 'IT00000000000000000000001004170008','IT00000000000000000000001004170009', 'IT00000000000000000000001004170010')
    """

    miniSubs = execute_query(query_minisubs)
    subs = execute_query(query_subs)
    oto = execute_query(query_oto)

    total_miniSubs = miniSubs['id'].count()
    total_subs = subs['id'].count()
    total_oto = oto['quantity'].sum()

    percentage_miniSubs = round(total_miniSubs/(total_subs+total_miniSubs)*100, 2)
    percentage_oto = round(total_oto/(total_subs+total_miniSubs+total_oto)*100, 2)

    subs_array = [total_miniSubs, total_subs, f"{percentage_miniSubs}%", f"{percentage_oto}%"]

    return subs_array












