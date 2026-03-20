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

def refill(start_date, end_date):

    query_refill = f"""
        SELECT SI.itemId, SO.order_number, SO.status, SO.created_at, SO.is_first_order, SI.quantity, SO.units
        FROM bi.fact_orders SO
        JOIN bi.fact_customers CU ON CU.id = SO.customer_id
        RIGHT JOIN bi.fact_sales_order_items SI 
            ON SI.salesOrderId = SO.id
        WHERE SI.category in ('IG00000000000000000000000000000044', 'IG00000000000000000000000000000043', 'IG00000000000000000000000000000041')
        AND SO.created_at > "{start_date}" AND SO.created_at < "{end_date}"
        AND SO.status <> "CANCELLED"
        AND SO.recurrent = false
        GROUP BY SO.id
        ORDER BY SO.created_at DESC;
    """

    query_total_oto_beard = f"""
        SELECT distinct
            fo.order_number, 
            fo.created_at, 
            so.category, 
            fo.is_first_order
        FROM bi.fact_orders fo
        JOIN bi.fact_sales_order_items so ON fo.id = so.salesOrderId
        where fo.status <> "CANCELLED"
        and fo.created_at > "{start_date}" AND fo.created_at < "{end_date}"
        and fo.order_plan = "OTO"
        and so.category in ("IG00000000000000000000000000000029");
    """

    rt = execute_query(query_refill)
    otob = execute_query(query_total_oto_beard)

    total_orders = rt['order_number'].count()
    new_customers = rt['is_first_order'].sum()
    units = rt['quantity'].sum() 
    oto_orders_beard = otob['order_number'].nunique()
    oto_orders_new_beard = otob['is_first_order'].sum()

    refillList = [total_orders, new_customers, units, oto_orders_beard, oto_orders_new_beard]

    return refillList




