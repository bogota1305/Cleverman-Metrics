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

def hear(start_date, end_date):

    query_total_orders_new = f"""
        SELECT distinct
            fo.order_number, 
            fo.created_at
        FROM bi.fact_orders fo
        where fo.status <> "CANCELLED"
        and fo.created_at BETWEEN "{start_date}" AND "{end_date}"
        and fo.is_first_order = 1;
    """

    query_hear_answered = f"""
        SELECT * FROM prod_ecommerce.customer_acquisition_source c 
        WHERE c.createdAt > "{start_date}" AND c.createdAt < "{end_date}"
        GROUP BY c.customerId
        ORDER BY c.createdAt ;
    """

    query_hear_options = f"""
        SELECT * FROM prod_ecommerce.customer_acquisition_source c 
        WHERE c.createdAt > "{start_date}" AND c.createdAt < "{end_date}"
        ORDER BY c.source;
    """

    ha = execute_query(query_hear_answered)
    ton = execute_query(query_total_orders_new)
    ho = execute_query(query_hear_options)

    total_orders_new = ton['order_number'].count()
    hear_answered = ha['id'].count()

    google = (ho['source'] == 'GOOGLE').sum()
    meta = (ho['source'] == 'FACEBOOK_INSTAGRAM').sum()
    word = (ho['source'] == 'WORD_OF_MOUTH').sum()
    amazon = (ho['source'] == 'AMAZON').sum()
    reddit = (ho['source'] == 'REDDIT').sum()
    chatgp = (ho['source'] == 'CHATGPT').sum()
    tiktok = (ho['source'] == 'TIKTOK').sum()
    other = (ho['source'] == 'OTHER').sum()

    hearTotalList = [total_orders_new, hear_answered]
    optionsList = [[google], [meta], [word], [amazon], [reddit], [chatgp], [tiktok], [other]] 

    return hearTotalList, optionsList







