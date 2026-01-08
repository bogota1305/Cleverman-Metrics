import pandas as pd
from modules.database_queries import execute_query

# Lista de productos (itemId y nombre)
productos = {
    'IT00000000000000000000000000000061': '45ml Colorant - Light Auburn',
    'IT00000000000000000000000000000060': '45ml Colorant - Dark Auburn',
    'IT00000000000000000000000000000059': '45ml Colorant - Lightest Blond',
    'IT00000000000000000000000000000058': '45ml Colorant - Warm - Light Blond',
    'IT00000000000000000000000000000057': '45ml Colorant - Cool - Light Blond',
    'IT00000000000000000000000000000056': '45ml Colorant - Warm-Medium Blond',
    'IT00000000000000000000000000000055': '45ml Colorant - Cool-Medium Blond',
    'IT00000000000000000000000000000054': '45ml Colorant - Warm-Dark Blond',
    'IT00000000000000000000000000000053': '45ml Colorant - Dark Blond',
    'IT00000000000000000000000000000052': '45ml Colorant: Light Brown',
    'IT00000000000000000000000000000051': '45ml Colorant - Medium-Light Brown',
    'IT00000000000000000000000000000050': '45ml Colorant - Medium Brown',
    'IT00000000000000000000000000000049': '45ml Colorant - Medium-Dark Brown',
    'IT00000000000000000000000000000048': '45ml Colorant - Dark Brown',
    'IT00000000000000000000000000000047': '45ml Colorant - Soft-Black',
    'IT00000000000000000000000000000046': '45ml Colorant - Black',
    'IT00000000000000000000000000000045': '45ml Colorant - Jet-Black',
    'IT00000000000000000000000000000021': '30ml Colorant - Light Auburn',
    'IT00000000000000000000000000000020': '30ml Colorant - Dark Auburn',
    'IT00000000000000000000000000000019': '30ml Colorant - Lightest Blond',
    'IT00000000000000000000000000000018': '30ml Colorant - Warm - Light Blond',
    'IT00000000000000000000000000000017': '30ml Colorant - Cool - Light Blond',
    'IT00000000000000000000000000000016': '30ml Colorant - Warm-Medium Blond',
    'IT00000000000000000000000000000015': '30ml Colorant - Cool-Medium Blond',
    'IT00000000000000000000000000000014': '30ml Colorant - Warm-Dark Blond',
    'IT00000000000000000000000000000013': '30ml Colorant - Dark Blond',
    'IT00000000000000000000000000000012': '30ml Colorant: Light Brown',
    'IT00000000000000000000000000000011': '30ml Colorant - Medium-Light Brown',
    'IT00000000000000000000000000000002': '30ml Colorant - Medium Brown',
    'IT00000000000000000000000000000010': '30ml Colorant - Medium-Dark Brown',
    'IT00000000000000000000000000000009': '30ml Colorant - Dark Brown',
    'IT00000000000000000000000000000008': '30ml Colorant - Soft-Black',
    'IT00000000000000000000000000000007': '30ml Colorant - Black',
    'IT00000000000000000000000000000006': '30ml Colorant - Jet-Black'
}

# Ejecutar la consulta SQL
query = """
SELECT 
    c.subscriptionId AS subscription_id,
    c.createdAt,
    sv.legacy_category,
    si.itemId
FROM 
    prod_sales_and_subscriptions.cancellations c
INNER JOIN 
    prod_sales_and_subscriptions.subscriptions_view sv
    ON c.subscriptionId = sv.subscription_id
INNER JOIN 
    prod_sales_and_subscriptions.subscription_items si
    ON c.subscriptionId = si.subscriptionId
WHERE 
    si.itemId IN ({})
    AND sv.status = 'CANCELLED'
""".format(", ".join(f"'{item}'" for item in productos.keys()))

# Obtener los datos
data = execute_query(query)

# Convertir los datos en un DataFrame de pandas
df = pd.DataFrame(data, columns=['subscription_id', 'createdAt', 'legacy_category', 'itemId'])

# Mapear los itemId a nombres de productos
df['producto'] = df['itemId'].map(productos)

# Convertir la columna createdAt a tipo datetime
df['createdAt'] = pd.to_datetime(df['createdAt'])

# Extraer la fecha (sin la hora)
df['fecha'] = df['createdAt'].dt.date

# Agrupar por fecha y producto, y contar las cancelaciones
cancelaciones_por_dia = df.groupby(['fecha', 'producto']).size().unstack(fill_value=0)

# Resetear el índice para que la fecha sea una columna
cancelaciones_por_dia = cancelaciones_por_dia.reset_index()

# Renombrar la columna de fechas a "date"
cancelaciones_por_dia = cancelaciones_por_dia.rename(columns={'fecha': 'date'})

# Calcular los totales por producto (suma de cada columna)
totales = cancelaciones_por_dia.iloc[:, 1:].sum()  # Ignorar la columna 'date' al calcular los totales

# Convertir la Serie de totales en un DataFrame
totales_df = pd.DataFrame([totales], columns=totales.index)

# Agregar la columna 'date' con el valor 'Total'
totales_df.insert(0, 'date', 'Total')

# Concatenar el DataFrame original con la fila de totales
cancelaciones_por_dia = pd.concat([cancelaciones_por_dia, totales_df], ignore_index=True)

# Calcular los porcentajes históricos
porcentajes_historicos = (totales / totales.sum()).to_frame().T
porcentajes_historicos['date'] = 'Porcentaje Histórico'

# Concatenar la fila de porcentajes históricos al DataFrame
cancelaciones_por_dia = pd.concat([cancelaciones_por_dia, porcentajes_historicos], ignore_index=True)

# Crear una nueva columna para el año y el mes
df['año_mes'] = df['createdAt'].dt.to_period('M')

# Agrupar por año, mes y producto, y contar las cancelaciones
cancelaciones_por_mes = df.groupby(['año_mes', 'producto']).size().unstack(fill_value=0)

# Resetear el índice para que el año y el mes sean columnas
cancelaciones_por_mes = cancelaciones_por_mes.reset_index()

# Renombrar la columna de año y mes a "date"
cancelaciones_por_mes = cancelaciones_por_mes.rename(columns={'año_mes': 'date'})

# Crear un archivo Excel con todas las hojas
with pd.ExcelWriter("colorCancelations.xlsx", engine='xlsxwriter') as writer:
    # Hoja General (cancelaciones por día)
    cancelaciones_por_dia.to_excel(writer, sheet_name='General', index=False)

    # Ajustar el tamaño de las celdas al texto
    workbook = writer.book
    worksheet = writer.sheets['General']
    gray_format_total = workbook.add_format({'bg_color': '#D3D3D3'})  # Formato gris claro
    gray_format_percentage = workbook.add_format({'bg_color': '#D3D3D3', 'num_format': '0.00%'})  # Formato gris claro
    percent_format = workbook.add_format({'num_format': '0.00%'})  # Formato de porcentaje

     # Aplicar formato gris claro a la fila de totales (penúltima fila)
    worksheet.set_row(cancelaciones_por_dia.shape[0] - 1, cell_format=gray_format_total)

    # Aplicar formato gris claro a la fila de porcentajes (última fila)
    worksheet.set_row(cancelaciones_por_dia.shape[0], cell_format=gray_format_percentage)

    for i, col in enumerate(cancelaciones_por_dia.columns):
        max_len = max(cancelaciones_por_dia[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)

    # Hojas por año (totales y porcentajes)
    for año, grupo in cancelaciones_por_mes.groupby(cancelaciones_por_mes['date'].dt.year):
        # Hoja de totales por mes
        grupo_totales = grupo.copy()
        grupo_totales['date'] = grupo_totales['date'].dt.strftime('%B')  # Convertir a nombre del mes

        # Calcular los totales por columna
        totales_mes = grupo_totales.iloc[:, 1:].sum()
        totales_mes['date'] = 'Total'
        grupo_totales = pd.concat([grupo_totales, pd.DataFrame([totales_mes])], ignore_index=True)

        grupo_totales.to_excel(writer, sheet_name=f"{año} - Totales", index=False)

        # Ajustar el tamaño de las celdas al texto
        worksheet = writer.sheets[f"{año} - Totales"]
        for i, col in enumerate(grupo_totales.columns):
            max_len = max(grupo_totales[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

         # Aplicar formato gris claro a la fila de totales (última fila)
        worksheet.set_row(grupo_totales.shape[0], cell_format=gray_format_total)

        # Hoja de porcentajes por mes
        grupo_porcentajes = grupo.copy()
        grupo_porcentajes['date'] = grupo_porcentajes['date'].dt.strftime('%B')  # Convertir a nombre del mes

        # Calcular los porcentajes (dividir entre el total mensual)
        grupo_porcentajes.iloc[:, 1:] = (grupo_porcentajes.iloc[:, 1:].div(grupo_porcentajes.iloc[:, 1:].sum(axis=1), axis=0))

        # Calcular los porcentajes totales (dividir entre el total general)
        totales_generales = grupo.iloc[:, 1:].sum()  # Totales generales por producto
        totales_porcentajes = (totales_generales / totales_generales.sum()).to_frame().T  # Porcentajes totales
        totales_porcentajes['date'] = 'Total'
        grupo_porcentajes = pd.concat([grupo_porcentajes, totales_porcentajes], ignore_index=True)

        grupo_porcentajes.to_excel(writer, sheet_name=f"{año} - Porcentajes", index=False)

        # Ajustar el tamaño de las celdas al texto y formatear porcentajes
        worksheet = writer.sheets[f"{año} - Porcentajes"]
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'num_format': '0.00%'})  # Formato rojo claro con porcentaje
        green_format = workbook.add_format({'bg_color': '#C0E5BB', 'num_format': '0.00%'})  # Formato rojo claro con porcentaje
        for i, col in enumerate(grupo_porcentajes.columns):
            max_len = max(grupo_porcentajes[col].astype(str).map(len).max(), len(col)) + 2
            if col != 'date':  # Aplicar formato de porcentaje a todas las columnas excepto 'date'
                worksheet.set_column(i, i, max_len, percent_format)
            else:
                worksheet.set_column(i, i, max_len)

        # Resaltar celdas donde el porcentaje es mayor en un 3% al histórico
        for row in range(0, grupo_porcentajes.shape[0]):  # Ignorar la fila de encabezados
            for col in range(1, grupo_porcentajes.shape[1]):  # Ignorar la columna de fechas
                valor_mes = grupo_porcentajes.iat[row, col]
                valor_historico = porcentajes_historicos.iat[0, col - 1]  # Restar 1 porque no hay columna 'date'
                if (1 - (valor_mes/valor_historico)) > 0.4 : 
                    worksheet.write(row+1, col, valor_mes, green_format)
                if (1 - (valor_mes/valor_historico)) < (-0.4) : 
                    worksheet.write(row+1, col, valor_mes, red_format)

        # Aplicar formato gris claro a la fila de totales (última fila)
        worksheet.set_row(grupo_porcentajes.shape[0], cell_format=gray_format_percentage)

print("Archivo Excel creado correctamente: colorCancelations.xlsx")