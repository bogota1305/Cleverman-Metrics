
import calendar
import datetime
from block_payments import get_blocked_payments
from exceptedRenewals import get_expected_renewals
from fullContol import fullControl
from ga4Funnels import get_funnel
from modules.date_selector import open_date_selector
from orders import get_orders
from payments import get_payments
from realRenewalFrecuency import realRenewalFrequency
from renewalsAndNoRecurrents import get_sales
from report import anotar_datos_excel, seleccionar_donde_almacenar, seleccionar_tipo_de_reporte
from selectFiles import seleccionar_archivos_para_casos, seleccionar_archivos_stripe
from subscriptions import subs
from uploadCloud import upload_to_drive, upload_to_dropbox

funnels_report = True
database_report = True
stripe_block_payments = True
folder_name = 'funnels'
columna = 16
dropbox_var = False
drive_var = False

funnels_report, database_report, stripe_block_payments = seleccionar_tipo_de_reporte()

if funnels_report:
    archivos = seleccionar_archivos_para_casos()
    
    if(database_report == False):
        dropbox_var, drive_var = seleccionar_donde_almacenar()

if database_report:
    start_date, end_date, folder_name, all_var, orders_var, unique_orders_var, sales_var, payment_errors_var, expected_renewals_var, frequency_var, full_control_var, subs_var = open_date_selector()

    dropbox_var, drive_var = seleccionar_donde_almacenar()

    values, items, urls = get_orders(start_date, end_date, folder_name, unique_orders_var, dropbox_var, drive_var)

    date_obj = datetime.datetime.strptime(start_date, '%Y-%m-%d')
    month_number = date_obj.month
    actualMonth = calendar.month_name[month_number]

    anotar_datos_excel(values, columna, 45+12, False, actualMonth, True)
    anotar_datos_excel(items, columna, 55+12, False, actualMonth)

    if(dropbox_var or drive_var):
        anotar_datos_excel(urls, columna, 45+12, True, actualMonth)
        anotar_datos_excel(urls, columna, 55+12, True, actualMonth)

    if(sales_var == 1):
        total_sales, urls = get_sales(start_date, end_date, folder_name, dropbox_var, drive_var)
        anotar_datos_excel(total_sales, columna, 67+12, False, actualMonth) 

        if(dropbox_var or drive_var):
            anotar_datos_excel(urls, columna, 67+12, True, actualMonth)

    if(payment_errors_var == 1):
        total_payments, urls = get_payments(start_date, end_date, folder_name, dropbox_var, drive_var) 
        anotar_datos_excel(total_payments, columna, 72+12, False, actualMonth)

        if(dropbox_var or drive_var):
            anotar_datos_excel(urls, columna, 72+12, True, actualMonth)

    if(expected_renewals_var == 1):
        total_expected_renewals = get_expected_renewals(start_date, end_date, folder_name) 
    
    if(frequency_var == 1):
        realRenewalFrequency(start_date, end_date, folder_name) 
    
    if(full_control_var == 1):
        percentageFC, newSubsFC, totalSubsFC, renewalFC= fullControl(start_date, end_date)
        anotar_datos_excel(newSubsFC, columna, 77+12, False, actualMonth)
        anotar_datos_excel(percentageFC, columna, 80+12, False, actualMonth)
        anotar_datos_excel(totalSubsFC, columna, 79+12, False, actualMonth)
        anotar_datos_excel(renewalFC, columna, 78+12, False, actualMonth)

    if(subs_var == 1):
        subsPercentage = subs(start_date, end_date)
        anotar_datos_excel(subsPercentage, columna, 88+12, False, actualMonth)

if funnels_report:
    if archivos['Customized Kit - Funnel'] != None:
        get_funnel(archivos['Customized Kit - Funnel'], 'Customized Kit - Funnel.xlsx', columna, 5, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['All In One - Funnel'] != None:
        get_funnel(archivos['All In One - Funnel'], 'All In One - Funnel.xlsx', columna, 12, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['Shop - Funnel'] != None:
        get_funnel(archivos['Shop - Funnel'], 'Shop - Funnel.xlsx', columna, 17, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['My Account - Funnel'] != None:
        get_funnel(archivos['My Account - Funnel'], 'My Account - Funnel.xlsx', columna, 22, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['Buy Again - Funnel'] != None:
        get_funnel(archivos['Buy Again - Funnel'], 'Buy Again - Funnel.xlsx', columna, 26, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['Buy Again Reactivate - Funnel'] != None:
        get_funnel(archivos['Buy Again Reactivate - Funnel'], 'Buy Again Reactivate - Funnel.xlsx', columna, 31, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['Buy Again Without Sub - Funnel'] != None:
        get_funnel(archivos['Buy Again Without Sub - Funnel'], 'Buy Again Without Sub - Funnel.xlsx', columna, 37, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['My Subscriptions - Funnel'] != None:
        get_funnel(archivos['My Subscriptions - Funnel'], 'My Subscriptions - Funnel.xlsx', columna, 31+12, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['NPD mail - Funnel'] != None:
        get_funnel(archivos['NPD mail - Funnel'], 'NPD mail - Funnel.xlsx', columna, 37+12, folder_name, dropbox_var, drive_var, actualMonth)

    if archivos['NPD account - Funnel'] != None:
        get_funnel(archivos['NPD account - Funnel'], 'NPD account - Funnel.xlsx', columna, 41+12, folder_name, dropbox_var, drive_var, actualMonth)

    indiceLandingsBeard = 107
    indiceLandingsHair = 150
    indiceLandingsOther = 187
    avanceLandings = 6

    # Orden exacto según tu lista
    beard_keys = [
        "Beard - JetBlack",
        "Beard - AfricanAmerican",
        "Beard - Black",
        "Beard - Blond",
        "Beard - Red",
        "Beard - Brown",
        "Beard - MediumDarkBrown",
    ]

    hair_keys = [
        "Hair - AfricanAmerican",
        "Hair - Red",
        "Hair - Black",
        "Hair - Brown",
        "Hair - LightBrown",
        "Hair - Blond",
    ]

    other_keys = [
        "My instructions",
        "Inmediate coverage",
        "Referral",
        "Grey hair color touch up",
        "Salt and peper",
        "Full coverage",
        "Customized beard",
        "Best beard color",
        "Best hair color one time",
        "Best hair color sub",
        "Customized beard sub",
    ]

    # Helper para no repetir lógica
    def process_funnels(keys, start_index):
        idx = start_index
        for k in keys:
            if archivos.get(k) is not None:
                get_funnel(
                    archivos[k],
                    f"{k}.xlsx",
                    columna,
                    idx,
                    folder_name,
                    dropbox_var,
                    drive_var,
                    actualMonth
                )
                idx += avanceLandings
        return idx

    # Ejecutar por secciones
    indiceLandingsBeard = process_funnels(beard_keys, indiceLandingsBeard)
    indiceLandingsHair  = process_funnels(hair_keys,  indiceLandingsHair)
    indiceLandingsOther = process_funnels(other_keys, indiceLandingsOther)


if stripe_block_payments:
    archivos = seleccionar_archivos_stripe() 
    if archivos['Blocked Payments'] != None and archivos['All Payments'] != None:
        get_blocked_payments(archivos['Blocked Payments'], archivos['All Payments'], 'Blocked payments', 'Stripe data')

nuevo_archivo = f'Monthly Report {actualMonth}.xlsx'

if(dropbox_var):
    upload_to_dropbox(nuevo_archivo, dropbox_path=f"/MyReports/{folder_name}/{nuevo_archivo}")
if(drive_var):  
    upload_to_drive(nuevo_archivo, folder_id="1F1VZxlp5IxkQEo4WD0Bt8VEJZ28OhGut")    
