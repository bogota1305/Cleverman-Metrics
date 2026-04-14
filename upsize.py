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

def upsize(start_date, end_date):

    query_scrub_upsized_orders = f"""
        SELECT SI.itemId, SO.order_number, SO.status, SO.created_at, SO.is_first_order, SI.quantity, SO.units
        FROM bi.fact_orders SO
        JOIN bi.fact_customers CU ON CU.id = SO.customer_id
        RIGHT JOIN bi.fact_sales_order_items SI 
            ON SI.salesOrderId = SO.id
        WHERE SI.itemId in ("IT00000000000000000000000000000110", "IT00000000000000000000000000000111", "IT00000000000000000000000000000112")
        AND SO.created_at > "{start_date}" AND SO.created_at < "{end_date}"
        AND SO.status <> "CANCELLED"
        GROUP BY SO.id
        ORDER BY SO.created_at DESC;
    """

    query_wipes_total_orders = f"""
        SELECT SI.itemId, SO.order_number, SO.status, SO.created_at, SO.is_first_order, SI.quantity, SO.units
        FROM bi.fact_orders SO
        JOIN bi.fact_customers CU ON CU.id = SO.customer_id
        RIGHT JOIN bi.fact_sales_order_items SI 
            ON SI.salesOrderId = SO.id
        WHERE SI.itemId = "IT00000000000000000000000000000115"
        AND SO.created_at > "{start_date}" AND SO.created_at < "{end_date}"
        AND SO.status <> "CANCELLED"
        GROUP BY SO.id
        ORDER BY SO.created_at DESC;
    """

    query_shampoo_total_orders = f"""
        SELECT SI.itemId, SO.order_number, SO.status, SO.created_at, SO.is_first_order, SI.quantity, SO.units
        FROM bi.fact_orders SO
        JOIN bi.fact_customers CU ON CU.id = SO.customer_id
        RIGHT JOIN bi.fact_sales_order_items SI 
            ON SI.salesOrderId = SO.id
        WHERE SI.itemId in ("IT00000000000000000000000000000246", "IT00000000000000000000000000000245")
        AND SO.created_at > "{start_date}" AND SO.created_at < "{end_date}"
        AND SO.status <> "CANCELLED"
        GROUP BY SO.id
        ORDER BY SO.created_at DESC;
    """

    query_conditioner_total_orders = f"""
        SELECT SI.itemId, SO.order_number, SO.status, SO.created_at, SO.is_first_order, SI.quantity, SO.units
        FROM bi.fact_orders SO
        JOIN bi.fact_customers CU ON CU.id = SO.customer_id
        RIGHT JOIN bi.fact_sales_order_items SI 
            ON SI.salesOrderId = SO.id
        WHERE SI.itemId = "IT00000000000000000000000000000248"
        AND SO.created_at > "{start_date}" AND SO.created_at < "{end_date}"
        AND SO.status <> "CANCELLED"
        GROUP BY SO.id
        ORDER BY SO.created_at DESC;
    """
    query_total_oto_orders = f"""
        SELECT distinct
            fo.order_number, 
            fo.created_at, 
            so.category 
        FROM bi.fact_orders fo
        JOIN bi.fact_sales_order_items so ON fo.id = so.salesOrderId
        where fo.status <> "CANCELLED"
        and fo.created_at > "{start_date}" AND fo.created_at < "{end_date}"
        and fo.order_plan = "OTO"
        and so.category in ("IG00000000000000000000000000000028", "IG00000000000000000000000000000029");
    """

    usto = execute_query(query_scrub_upsized_orders)

    uwto = execute_query(query_wipes_total_orders)
    
    ushoto = execute_query(query_shampoo_total_orders)
    
    ucto = execute_query(query_conditioner_total_orders)

    to = execute_query(query_total_oto_orders)

    beard_total_orders = (to['category'] == 'IG00000000000000000000000000000029').sum()
    scrub_upsized_orders = usto['order_number'].nunique()
    scrub_sachets = usto['quantity'].sum()

    wipes_total_orders = to['order_number'].nunique()
    wipes_upsized_orders = uwto['order_number'].nunique()
    wipes_sachets = uwto['quantity'].sum()

    hair_total_orders = (to['category'] == 'IG00000000000000000000000000000028').sum()
 
    shampoo_upsized_orders = ushoto['order_number'].nunique()
    shampoo_sachets = ushoto['quantity'].sum()

    conditioner_upsized_orders = ucto['order_number'].nunique()
    conditioner_sachets = ucto['quantity'].sum()

    total_orders = wipes_total_orders
    upsized_orders = scrub_upsized_orders + wipes_upsized_orders + shampoo_upsized_orders + conditioner_upsized_orders
    sachets = scrub_sachets + wipes_sachets + shampoo_sachets + conditioner_sachets

    bearScrubList = [beard_total_orders, scrub_upsized_orders, scrub_sachets]
    wipesList = [wipes_total_orders, wipes_upsized_orders, wipes_sachets]
    totalList = [total_orders, upsized_orders, sachets]
    shampooList = [hair_total_orders, shampoo_upsized_orders, shampoo_sachets]
    conditionerList = [hair_total_orders, conditioner_upsized_orders, conditioner_sachets]

    return bearScrubList, wipesList, totalList, shampooList, conditionerList








