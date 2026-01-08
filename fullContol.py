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

def fullControl(start_date, end_date):

    query_total = f"""
    SELECT distinct id
    FROM bi.fact_subscriptions 
    where sms_renewal = 1
    AND status != 'CANCELLED';
    """

    query_new = f"""
    SELECT
        v.subscription_id,
        s.status,
        DATE(CONVERT_TZ(v.first_date_sms_renewal_true, 'UTC', 'America/Los_Angeles')) AS full_control_starting_date,
        COUNT(r.id) AS sucess_renewal_count_after_first,
        GROUP_CONCAT(
            DATE(CONVERT_TZ(r.createdAt, 'UTC', 'America/Los_Angeles'))
            ORDER BY r.createdAt ASC
            SEPARATOR ','
        ) AS success_renewal_dates_after_first
    FROM
        prod_sales_and_subscriptions.first_sms_renewal_versions v
    JOIN
    prod_sales_and_subscriptions.renewals r
    ON v.subscription_id = r.subscriptionId
    JOIN
        prod_sales_and_subscriptions.subscriptions s
        ON s.id = v.subscription_id
    WHERE v.first_date_sms_renewal_true > "{start_date}" AND v.first_date_sms_renewal_true < "{end_date}"
    GROUP BY
        v.subscription_id,
        DATE(CONVERT_TZ(v.first_date_sms_renewal_true, 'UTC', 'America/Los_Angeles'));
    """

    query_renewal = f"""
    SELECT
        v.subscription_id,
        s.status,
        DATE(CONVERT_TZ(v.first_date_sms_renewal_true, 'UTC', 'America/Los_Angeles')) AS full_control_starting_date,
        COUNT(r.id) AS sucess_renewal_count_after_first,
        GROUP_CONCAT(
            DATE(CONVERT_TZ(r.createdAt, 'UTC', 'America/Los_Angeles'))
            ORDER BY r.createdAt ASC
            SEPARATOR ','
        ) AS success_renewal_dates_after_first
    FROM
        prod_sales_and_subscriptions.first_sms_renewal_versions v
    JOIN
        prod_sales_and_subscriptions.renewals r
        ON v.subscription_id = r.subscriptionId
        AND r.createdAt > DATE_ADD(v.first_date_sms_renewal_true, INTERVAL 1 DAY)
        AND r.status = 'SUCCESS'
        AND r.createdAt > "{start_date}"
        AND (v.last_date_sms_renewal_true is null OR r.createdAt < v.last_date_sms_renewal_true)
    JOIN
        prod_sales_and_subscriptions.subscriptions s
        ON s.id = v.subscription_id
    -- WHERE v.first_date_sms_renewal_true > "2024-01-01 08:00:00" -- AND v.first_date_sms_renewal_true < "2025-05-01 08:00:00"
    GROUP BY
        v.subscription_id,
        DATE(CONVERT_TZ(v.first_date_sms_renewal_true, 'UTC', 'America/Los_Angeles'));
    """

    fct = execute_query(query_total)
    fcn = execute_query(query_new)
    fcr = execute_query(query_renewal)

    total_subs = fct['id'].count()
    new_subs = fcn['subscription_id'].count()
    renewals = fcr['subscription_id'].count()

    percentage_new_subs = round(new_subs/total_subs*100, 2)

    percentage_array = [f"{percentage_new_subs}%"]
    new_subs_array = [new_subs] 
    total_subs_array = [total_subs] 
    renewal_subs_array = [renewals] 

    print(f"Porcentage: {percentage_new_subs} Nuevas {new_subs} Totales {total_subs} renovaciones {renewals}")

    return percentage_array, new_subs_array, total_subs_array, renewal_subs_array








