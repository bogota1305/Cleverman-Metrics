from tkinter import simpledialog
import pandas as pd
import matplotlib.pyplot as plt

from modules.excel_creator import save_dataframe_to_excel_ga4
from report import anotar_datos_excel

def get_funnel(ruta_archivo, nombre_salida, columna_inicio, fila_inicio, carpeta_salida, dropbox_var, drive_var, month =''):
    
    data = pd.read_csv(ruta_archivo, encoding='utf-8', skiprows=9)

    # Eliminar la columna "Ingreso Active users"
    data = data.drop(columns=['Ingreso Active users'], errors='ignore')

    # Ordenar las filas por la columna 'Day'
    data = data.sort_values(by='Day', ascending=True)

    # Convertir los valores numéricos en 'Day' a enteros para asegurar un orden correcto
    data['Day'] = pd.to_numeric(data['Day'], errors='coerce')

    # Reorganizar los datos para que los pasos sean filas y los días sean columnas
    pivot_days = data.set_index('Day').T

    # Crear una columna "Total" que sume los valores de todos los días
    pivot_days['Total'] = pivot_days.sum(axis=1)

    # Limpiar los nombres de los pasos
    pivot_days.index = pivot_days.index.str.replace(r'^\d+\.\s*', '', regex=True)

    # Identificar la primera fila como referencia para calcular los porcentajes
    reference_row = pivot_days.iloc[0]  # Seleccionar la primera fila
    reference_row_name = pivot_days.index[0]  # Guardar el nombre de la fila de referencia

    # Calcular la tabla de porcentajes basada en la primera fila
    percentages_table = pivot_days.div(reference_row) * 100

    # Renombrar las filas para diferenciar las tablas
    percentages_table.index = [f"{step} (%)" for step in percentages_table.index]

    # Concatenar la tabla original con la tabla de porcentajes
    final_table_with_percentages = pd.concat([pivot_days, percentages_table])

    # Separar las tablas de números y porcentajes
    counts_table = final_table_with_percentages[~final_table_with_percentages.index.str.contains(r'\(%\)', regex=True)]
    percentages_table = final_table_with_percentages[final_table_with_percentages.index.str.contains(r'\(%\)', regex=True)]

    # Excluir la primera fila de la tabla de porcentajes (que representa el 100%)
    percentages_table = percentages_table[percentages_table.index != f"{reference_row_name} (%)"]

    # Formatear la tabla de porcentajes como valores porcentuales
    percentages_table = percentages_table.applymap(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")

    # Concatenar las tablas con una fila de espacio en blanco entre ambas
    empty_row = pd.DataFrame([["" for _ in counts_table.columns]], columns=counts_table.columns)
    final_table_spaced = pd.concat([counts_table, empty_row, percentages_table])

    # Calcular los porcentajes relativos al paso anterior
    percentages_previous_step = pivot_days.copy() 

    for column in percentages_previous_step.columns:
        percentages_previous_step[column] = percentages_previous_step[column] / percentages_previous_step[column].shift(1) 
        percentages_previous_step[column] *= 100

    # Reemplazar valores indefinidos (NaN) y negativos con 0%
    percentages_previous_step = percentages_previous_step.fillna(0).clip(lower=0)

    # Formatear la tabla de porcentajes como valores porcentuales
    percentages_previous_step = percentages_previous_step.applymap(
        lambda x: f"{x:.2f}%" if pd.notnull(x) else "")

    # Renombrar las filas para indicar el paso anterior
    steps_with_previous = percentages_previous_step.index.tolist()
    steps_with_previous = [
        f"{steps_with_previous[i]} (vs '{steps_with_previous[i - 1]}')"
        if i > 0 else steps_with_previous[i] for i in range(len(steps_with_previous))
    ]
    percentages_previous_step.index = steps_with_previous

    # Excluir la primera fila ya que no tiene paso anterior
    percentages_previous_step = percentages_previous_step.iloc[1:]

    # Concatenar la nueva tabla con un espacio en blanco al final de las demás tablas
    empty_row_previous = pd.DataFrame(
        [["" for _ in pivot_days.columns]], columns=pivot_days.columns)
    final_table_spaced_with_previous = pd.concat(
        [final_table_spaced, empty_row, percentages_previous_step]
    )

    urls = save_dataframe_to_excel_ga4(percentages_table, percentages_previous_step, final_table_spaced_with_previous, nombre_salida, carpeta_salida, dropbox_var, drive_var)

    # Obtener los datos de la última columna de percentages_previous_step como una lista
    # Obtener los datos de la última columna de percentages_table como una lista
    datos = percentages_table.iloc[:, -1].tolist()

    # Obtener el valor de reference_row correspondiente a la última columna
    reference_value = reference_row.iloc[-1]

    # Insertar el valor de reference_row en la primera posición de la lista de datos
    datos.insert(0, f"{int(reference_value)} (100%)")

    # Llamar a la función anotar_datos_excel con los datos actualizados
    anotar_datos_excel(datos, columna_inicio, fila_inicio, False, month)
    
    if(dropbox_var or drive_var):
        anotar_datos_excel(urls, columna_inicio, fila_inicio, True, month)
        
    