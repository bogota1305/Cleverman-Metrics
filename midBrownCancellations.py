import pandas as pd
from modules.database_queries import execute_query
from datetime import datetime

# Consulta SQL para obtener las cancelaciones del item específico
query30ml = """
SELECT 
    sc.subscriptionId,
    sc.reason,
    sc.createdAt
FROM 
    prod_sales_and_subscriptions.cancellations sc
INNER JOIN
    prod_sales_and_subscriptions.subscription_items si
    ON sc.subscriptionId = si.subscriptionId
WHERE
    si.itemId = 'IT00000000000000000000000000000002'
    AND sc.createdAt > '2022-01-01' AND sc.createdAt < '2025-04-01'
"""

query45ml = """
SELECT 
    sc.subscriptionId,
    sc.reason,
    sc.createdAt
FROM 
    prod_sales_and_subscriptions.cancellations sc
INNER JOIN
    prod_sales_and_subscriptions.subscription_items si
    ON sc.subscriptionId = si.subscriptionId
WHERE
    si.itemId = 'IT00000000000000000000000000000050'
    AND sc.createdAt > '2022-01-01' AND sc.createdAt < '2025-04-01'
"""

queryNoMediumBrown = """
SELECT 
    sc.subscriptionId,
    sc.reason,
    sc.createdAt
FROM 
    prod_sales_and_subscriptions.cancellations sc
INNER JOIN
    prod_sales_and_subscriptions.subscription_items si
    ON sc.subscriptionId = si.subscriptionId
WHERE
    si.itemId NOT IN ('IT00000000000000000000000000000002', 'IT00000000000000000000000000000050')
    AND sc.createdAt > '2022-01-01' AND sc.createdAt < '2025-04-01'
"""

def mediumBrownCancellationReasons(query, fileName):
    # Obtener los datos
    data = execute_query(query)

    # Convertir a DataFrame
    df = pd.DataFrame(data, columns=['subscriptionId', 'reason', 'createdAt'])

    # Procesar las razones (tomar solo la parte antes del primer "->")
    def simplificar_razon(razon):
        if pd.isna(razon):
            return "Sin razón especificada"
        partes = razon.split("->")
        partes = razon.split("-")
        return partes[0].strip()

    df['razon_simplificada'] = df['reason'].apply(simplificar_razon)

    # Convertir createdAt a datetime y extraer año y mes
    df['createdAt'] = pd.to_datetime(df['createdAt'])
    df['año'] = df['createdAt'].dt.year
    df['mes'] = df['createdAt'].dt.month
    df['mes_nombre'] = df['createdAt'].dt.strftime('%B')

    # Crear el archivo Excel
    with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
        # Calcular porcentajes históricos (promedio de todos los años)
        pivot_historico = df.pivot_table(
            index='razon_simplificada',
            values='subscriptionId',
            aggfunc='count',
            fill_value=0
        )
        porcentaje_historico = (pivot_historico / pivot_historico.sum()).T
        
        # Hoja Histórico
        porcentaje_historico.to_excel(writer, sheet_name='Histórico', index=False)
        
        # Formatear hoja Histórico
        worksheet_historico = writer.sheets['Histórico']
        percent_format = writer.book.add_format({'num_format': '0.00%'})
        
        # Ajustar ancho de columnas en hoja Histórico
        for i, col in enumerate(porcentaje_historico.columns):
            max_len = max(
                porcentaje_historico[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet_historico.set_column(i, i, max_len, percent_format)
        
        # Agrupar por año
        for año, grupo_anual in df.groupby('año'):
            # Hoja de cantidades
            pivot_cantidades = grupo_anual.pivot_table(
                index='mes_nombre',
                columns='razon_simplificada',
                values='subscriptionId',
                aggfunc='count',
                fill_value=0
            ).reset_index()
            
            # Ordenar por mes cronológicamente
            meses_orden = ['January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December']
            pivot_cantidades['mes_nombre'] = pd.Categorical(
                pivot_cantidades['mes_nombre'],
                categories=meses_orden,
                ordered=True
            )
            pivot_cantidades = pivot_cantidades.sort_values('mes_nombre')
            
            # Agregar fila de totales (solo para la hoja de cantidades)
            totales = pivot_cantidades.iloc[:, 1:].sum().to_frame().T
            totales.insert(0, 'mes_nombre', 'Total')
            pivot_cantidades_con_totales = pd.concat([pivot_cantidades, totales], ignore_index=True)
            
            # Escribir hoja de cantidades
            pivot_cantidades_con_totales.to_excel(writer, sheet_name=f"{año} - Cantidades", index=False)
            
            # Formatear hoja de cantidades
            worksheet = writer.sheets[f"{año} - Cantidades"]
            for i, col in enumerate(pivot_cantidades_con_totales.columns):
                max_len = max(
                    pivot_cantidades_con_totales[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(i, i, max_len)
            
            # Hoja de porcentajes
            pivot_porcentajes = pivot_cantidades.copy()
            pivot_porcentajes.iloc[:, 1:] = pivot_porcentajes.iloc[:, 1:].div(
                pivot_porcentajes.iloc[:, 1:].sum(axis=1), axis=0
            )
            
            # Calcular promedios por razón
            promedios = pivot_porcentajes.iloc[:, 1:].mean().to_frame().T
            promedios.insert(0, 'mes_nombre', 'Promedio')
            pivot_porcentajes_con_promedios = pd.concat([pivot_porcentajes, promedios], ignore_index=True)
            
            # Escribir hoja de porcentajes
            pivot_porcentajes_con_promedios.to_excel(writer, sheet_name=f"{año} - Porcentajes", index=False)
            
            # Formatear hoja de porcentajes
            worksheet = writer.sheets[f"{año} - Porcentajes"]
            workbook = writer.book
            
            # Crear formatos
            percent_format = workbook.add_format({'num_format': '0.00%'})
            red_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#FFC7CE'})  # Rojo claro
            green_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#C0E5BB'})  # Verde claro
            
            # Aplicar formato y resaltar variaciones >40% respecto al histórico
            for row in range(1, len(pivot_porcentajes)+1):  # Excluir fila de promedios
                for col in range(1, len(pivot_porcentajes.columns)):
                    valor_mes = pivot_porcentajes_con_promedios.iat[row-1, col]
                    razon = pivot_porcentajes_con_promedios.columns[col]
                    
                    # Obtener valor histórico para esta razón
                    if razon in porcentaje_historico.columns:
                        valor_historico = porcentaje_historico[razon].values[0]
                        
                        # Calcular diferencia porcentual
                        if valor_historico > 0:
                            diferencia = (valor_mes - valor_historico) / valor_historico
                            
                            # Aplicar formato según diferencia
                            if diferencia > 0.4:  # >40% mayor que histórico
                                worksheet.write(row, col, valor_mes, red_format)
                            elif diferencia < -0.4:  # >40% menor que histórico
                                worksheet.write(row, col, valor_mes, green_format)
                            else:
                                worksheet.write(row, col, valor_mes, percent_format)
                        else:
                            worksheet.write(row, col, valor_mes, percent_format)
                    else:
                        worksheet.write(row, col, valor_mes, percent_format)
            
            # Formatear fila de promedios (sin resaltado)
            last_row = len(pivot_porcentajes_con_promedios)
            for col in range(1, len(pivot_porcentajes_con_promedios.columns)):
                worksheet.write(last_row, col, pivot_porcentajes_con_promedios.iat[-1, col], percent_format)
            
            # Ajustar ancho de columnas
            for i, col in enumerate(pivot_porcentajes_con_promedios.columns):
                max_len = max(
                    pivot_porcentajes_con_promedios[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(i, i, max_len)

    print(f"Archivo Excel creado correctamente: {fileName}")

# mediumBrownCancellationReasons(query30ml, "30ml_medium_brown_cancellation_reasons.xlsx")
# mediumBrownCancellationReasons(query45ml, "45ml_medium_brown_cancellation_reasons.xlsx")
mediumBrownCancellationReasons(queryNoMediumBrown, "cancellation_reasons_no_medium_brown.xlsx")