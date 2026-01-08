import pandas as pd
from modules.database_queries import execute_query
from modules.excel_creator import save_dataframe_to_excel
import numpy as np  

def get_blocked_payments(ruta_archivo_blocked, ruta_archivo_payments, nombre_salida, carpeta_salida):
    # Cargar los archivos CSV
    df = pd.read_csv(ruta_archivo_blocked)
    dp = pd.read_csv(ruta_archivo_payments)
    
    # Convertir la columna 'rule_decision_created' a tipo datetime
    df['rule_decision_created'] = pd.to_datetime(df['rule_decision_created'])

    # Extraer la fecha sin la hora
    df['date'] = df['rule_decision_created'].dt.date

    # Función para extraer el device_id de la descripción
    def extract_device_id(description):
        if 'device_id:' in description:
            return description.split('device_id:')[-1].split()[0]  # Extrae el device_id
        return None

    # Aplicar la función para crear una nueva columna 'device_id'
    df['device_id'] = df['description'].apply(extract_device_id)

    # Aplicar la función para crear una nueva columna 'initial_id'
    df['initial_id'] = df['description'].apply(lambda x: x.split()[0])
    dp['initial_id'] = dp['Description'].apply(lambda x: x.split()[0])

    # Crear un diccionario para mapear customer_id a device_id
    customer_to_device = {}

    # Asignar device_id basado en customer_id
    for index, row in df.iterrows():
        customer_id = row['customer_id']
        device_id = row['device_id']
        
        if pd.notna(device_id):  # Si hay un device_id en la descripción
            if customer_id in customer_to_device:  # Si el customer_id ya está asociado a un device_id
                None
            else:  # Si no está asociado, agregar al diccionario
                customer_to_device[customer_id] = device_id
        elif customer_id in customer_to_device:  # Si no hay device_id pero el customer_id ya está asociado
            df.at[index, 'device_id'] = customer_to_device[customer_id]  # Asignar el device_id existente
        else:  # Si no hay device_id ni customer_id asociado
            new_device_id = f"user_{row['payment_intent_id']}"  # Crear un nuevo device_id basado en payment_intent_id
            customer_to_device[customer_id] = new_device_id
            df.at[index, 'device_id'] = new_device_id

    # Crear un diccionario para mapear initial_id a Status en dp
    # Si al menos una fila con el mismo initial_id tiene Status = "Paid", se marca como "Paid"
    id_to_status = dp.groupby('initial_id')['Status'].apply(lambda x: 'Paid' if 'Paid' in x.values else x.iloc[0]).to_dict()

    # Función para verificar si un pago está resuelto
    def is_payment_resolved(initial_id):
        return id_to_status.get(initial_id) == 'Paid'

    # Aplicar la función para crear una nueva columna 'is_resolved'
    df['is_resolved'] = df['initial_id'].apply(is_payment_resolved)

    # 1. Calcular real_block_payments (número de grupos únicos de 'payment_intent_id' por día)
    real_block_payments = df.groupby(['date', 'payment_intent_id']).size().reset_index().groupby('date').size().reset_index(name='real_block_payments')

    # 2. Calcular real_amount_blocked (suma del primer 'amount' de cada grupo de 'payment_intent_id' por día)
    real_amount_blocked = df.drop_duplicates(subset=['date', 'payment_intent_id'], keep='first').groupby('date')['amount'].sum().reset_index(name='real_amount_blocked')

    # 3. Calcular blocked_users (número de device_id únicos por día)
    # Si no hay device_id, usar customer_id como respaldo
    df['user_identifier'] = df['device_id'].fillna(df['customer_id'])
    blocked_users = df.groupby('date')['user_identifier'].nunique().reset_index(name='blocked_users')

    # 4. Calcular blocked_payments_resolved (número de pagos resueltos por día, contando solo uno por grupo de payment_intent_id)
    # Primero, filtramos los pagos resueltos
    resolved_payments = df[df['is_resolved']]
    # Luego, eliminamos duplicados por payment_intent_id, manteniendo solo el primero
    resolved_payments_unique = resolved_payments.drop_duplicates(subset=['date', 'payment_intent_id'], keep='first')
    # Finalmente, contamos por día
    blocked_payments_resolved = resolved_payments_unique.groupby('date').size().reset_index(name='blocked_payments_resolved')

    # 5. Calcular amount_resolved (suma de los valores de 'Amount' en dp para los pagos resueltos, contando solo uno por grupo de payment_intent_id)
    # Primero, filtramos los pagos resueltos en dp
    dp_resolved = dp[dp['Status'] == 'Paid']
    # Luego, eliminamos duplicados por initial_id, manteniendo solo el primero
    dp_resolved_unique = dp_resolved.drop_duplicates(subset=['initial_id'], keep='first')
    # Mapeamos initial_id a Amount en dp
    id_to_amount = dict(zip(dp_resolved_unique['initial_id'], dp_resolved_unique['Amount']))
    # Aplicamos el mapeo a resolved_payments_unique para obtener el Amount resuelto
    resolved_payments_unique['amount_resolved'] = resolved_payments_unique['initial_id'].map(id_to_amount).fillna(0)
    # Sumamos por día
    amount_resolved = resolved_payments_unique.groupby('date')['amount_resolved'].sum().reset_index(name='amount_resolved')

    # 6. Calcular blocked_users_resolved (número de usuarios resueltos únicos por día)
    # Primero, filtramos los pagos resueltos
    resolved_payments = df[df['is_resolved']]
    # Luego, eliminamos duplicados por device_id, manteniendo solo el primero
    resolved_users_unique = resolved_payments.drop_duplicates(subset=['date', 'device_id'], keep='first')
    # Finalmente, contamos los usuarios resueltos únicos por día
    blocked_users_resolved = resolved_users_unique.groupby('date').size().reset_index(name='blocked_users_resolved')

    # 7. Calcular new_users_blocked (número de nuevos usuarios bloqueados por día)
    # Agrupamos por payment_intent_id y verificamos si no hay ningún device_id en el grupo
    new_users_blocked = (
        df.groupby(['date', 'payment_intent_id'])['description']
        .apply(lambda x: 1 if x.str.contains('device_id:').any() == False else 0)  # Contar 1 si no hay 'device_id:' en ninguna fila del grupo
        .reset_index(name='new_users_blocked')
        .groupby('date')['new_users_blocked']
        .sum()  # Sumamos por día
        .reset_index(name='new_users_blocked')
    )

    # 8. Crear una lista de usuarios bloqueados únicos
    usuarios_bloqueados_unicos = df.drop_duplicates(subset=['device_id'])

    # 9. Filtrar usuarios que no tienen device_id o cuyo device_id no es un correo válido
    usuarios_sin_device_id = usuarios_bloqueados_unicos[usuarios_bloqueados_unicos['device_id'].isna()]
    usuarios_con_device_id_no_valido = usuarios_bloqueados_unicos[
        (usuarios_bloqueados_unicos['device_id'].notna()) & 
        (~usuarios_bloqueados_unicos['device_id'].str.contains('@', na=False))
    ]

    # 10. Contar estos usuarios como nuevos
    new_users_blocked = (
        pd.concat([usuarios_sin_device_id, usuarios_con_device_id_no_valido])
        .groupby('date').size().reset_index(name='new_users_blocked')
    )

    # 11. Filtrar los usuarios restantes (con device_id válido)
    usuarios_con_device_id_valido = usuarios_bloqueados_unicos[
        (usuarios_bloqueados_unicos['device_id'].notna()) & 
        (usuarios_bloqueados_unicos['device_id'].str.contains('@', na=False))
    ]

    # 12. Realizar la consulta a la base de datos para obtener todos los correos de la tabla customers
    consulta_customers = """
        SELECT id, email
        FROM prod_sales_and_subscriptions.customers;
    """
    customers_df = pd.DataFrame(execute_query(consulta_customers), columns=['id', 'email'])

    # 13. Crear un conjunto de correos existentes en la base de datos
    existing_emails = set(customers_df['email'].dropna().unique())

    # 14. Verificar si los device_id de los usuarios restantes existen en la base de datos
    usuarios_con_device_id_valido['existe_en_base_de_datos'] = usuarios_con_device_id_valido['device_id'].isin(existing_emails)

    # 15. Filtrar usuarios que no existen en la base de datos y contarlos como nuevos
    usuarios_nuevos_por_no_existir = usuarios_con_device_id_valido[~usuarios_con_device_id_valido['existe_en_base_de_datos']]
    new_users_blocked = pd.concat([new_users_blocked, usuarios_nuevos_por_no_existir.groupby('date').size().reset_index(name='new_users_blocked')])
    new_users_blocked = new_users_blocked.groupby('date')['new_users_blocked'].sum().reset_index()

    # 16. Filtrar usuarios que existen en la base de datos
    usuarios_existentes = usuarios_con_device_id_valido[usuarios_con_device_id_valido['existe_en_base_de_datos']]

    # 17. Realizar la consulta a la base de datos para obtener los sales_orders de los usuarios existentes
    if len(usuarios_existentes) > 0:
        # Obtener los IDs de los usuarios existentes
        ids_usuarios_existentes = customers_df[customers_df['email'].isin(usuarios_existentes['device_id'])]['id'].unique()
        if len(ids_usuarios_existentes) > 0:
            # Asegúrate de que los valores de customerId estén entre comillas simples
            consulta_sales_orders = f"""
                SELECT customerId, createdAt
                FROM prod_sales_and_subscriptions.sales_orders
                WHERE customerId IN ({','.join([f"'{id}'" for id in ids_usuarios_existentes])});
            """
            sales_orders_df = pd.DataFrame(execute_query(consulta_sales_orders), columns=['customerId', 'createdAt'])
            sales_orders_df['createdAt'] = pd.to_datetime(sales_orders_df['createdAt'])
        else:
            sales_orders_df = pd.DataFrame(columns=['customerId', 'createdAt'])  # Si no hay usuarios, crear un DataFrame vacío
    else:
        sales_orders_df = pd.DataFrame(columns=['customerId', 'createdAt'])  # Si no hay usuarios, crear un DataFrame vacío

    # 18. Para cada usuario existente, verificar si tiene algún sales_order creado al menos un día antes del pago bloqueado
    for index, row in usuarios_existentes.iterrows():
        usuario_id = customers_df[customers_df['email'] == row['device_id']]['id'].values[0]
        fecha_pago_bloqueado = row['rule_decision_created']
        sales_orders_usuario = sales_orders_df[sales_orders_df['customerId'] == usuario_id]
        
        # Verificar si hay algún sales_order creado al menos un día antes del pago bloqueado
        if not any(sales_orders_usuario['createdAt'] < (fecha_pago_bloqueado - pd.Timedelta(days=1))):
            # Si no tiene órdenes anteriores, el usuario se cuenta como nuevo
            new_users_blocked.loc[new_users_blocked['date'] == row['date'], 'new_users_blocked'] += 1

    # 19. Calcular new_users_blocked_resolved (número de usuarios nuevos resueltos)
    # Combinar new_users_blocked y blocked_users_resolved para obtener new_users_blocked_resolved
    new_users_blocked_resolved = pd.merge(new_users_blocked, blocked_users_resolved, on='date', how='left').fillna(0)
    new_users_blocked_resolved['new_users_blocked_resolved'] = new_users_blocked_resolved.apply(
        lambda row: min(row['new_users_blocked'], row['blocked_users_resolved']), axis=1
    )
    new_users_blocked_resolved = new_users_blocked_resolved[['date', 'new_users_blocked_resolved']]

    # 20. Combinar todos los resultados en una sola tabla
    resultado_final = pd.merge(real_block_payments, real_amount_blocked, on='date')
    resultado_final = pd.merge(resultado_final, blocked_users, on='date')
    resultado_final = pd.merge(resultado_final, blocked_payments_resolved, on='date', how='left').fillna(0)
    resultado_final = pd.merge(resultado_final, amount_resolved, on='date', how='left').fillna(0)
    resultado_final = pd.merge(resultado_final, blocked_users_resolved, on='date', how='left').fillna(0)
    resultado_final = pd.merge(resultado_final, new_users_blocked, on='date', how='left').fillna(0)
    resultado_final = pd.merge(resultado_final, new_users_blocked_resolved, on='date', how='left').fillna(0)

    # 21. Calcular percentage_new_users_blocked (porcentaje de nuevos usuarios bloqueados)
    # Obtener la primera y última fecha del archivo df
    fecha_inicio = df['date'].min().strftime('%Y-%m-%d %H:%M:%S')
    fecha_fin = (df['date'].max() + pd.Timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S')

    # Consulta SQL para obtener los usuarios nuevos
    consulta_sql = f"""
        SELECT *
        FROM bi.fact_orders
        WHERE created_at >= '{fecha_inicio}'
        AND created_at < '{fecha_fin}'
        AND is_first_order = 1;
    """
    # Ejecutar la consulta
    usuarios_nuevos = execute_query(consulta_sql)

    # Convertir el resultado de la consulta en un DataFrame
    usuarios_nuevos_df = pd.DataFrame(usuarios_nuevos, columns=['created_at', 'is_first_order'])

    # Convertir la columna 'created_at' a tipo datetime
    usuarios_nuevos_df['created_at'] = pd.to_datetime(usuarios_nuevos_df['created_at'])

    # Extraer la fecha sin la hora
    usuarios_nuevos_df['date'] = usuarios_nuevos_df['created_at'].dt.date

    # Contar el número de usuarios nuevos por día
    usuarios_nuevos_por_dia = usuarios_nuevos_df.groupby('date').size().reset_index(name='total_new_users')

    # Combinar con el DataFrame resultado_final
    resultado_final = pd.merge(resultado_final, usuarios_nuevos_por_dia, on='date', how='left').fillna(0)

    # 22. Calcular el porcentaje de nuevos usuarios bloqueados con la nueva fórmula
    resultado_final['percentage_new_users_blocked (%)'] = (
        resultado_final.apply(
            lambda row: 0 if row['total_new_users'] + row['new_users_blocked'] - row['new_users_blocked_resolved'] == 0
            else (row['new_users_blocked'] / (row['total_new_users'] + row['new_users_blocked'] - row['new_users_blocked_resolved'])) * 100,
            axis=1
        )
    ).round(2)  # Redondear a 2 decimales

    # Si new_users_blocked es 0, el porcentaje también debe ser 0
    resultado_final.loc[resultado_final['new_users_blocked'] == 0, 'percentage_new_users_blocked (%)'] = 0

    # Agregar una fila de totales
    total_new_users_blocked = resultado_final['new_users_blocked'].sum()
    total_total_new_users = resultado_final['total_new_users'].sum()
    total_blocked_users = resultado_final['blocked_users'].sum()
    total_blocked_users_resolved = resultado_final['blocked_users_resolved'].sum()
    total_new_users_blocked_resolved = resultado_final['new_users_blocked_resolved'].sum()

    # Calcular el total de usuarios nuevos ajustado
    total_total_new_users_ajustado = (
        total_total_new_users + 
        total_new_users_blocked - 
        total_new_users_blocked_resolved
    )

    totales = pd.DataFrame({
        'date': ['Total'],
        'real_block_payments': [resultado_final['real_block_payments'].sum()],
        'real_amount_blocked': [resultado_final['real_amount_blocked'].sum()],
        'blocked_users': [total_blocked_users],
        'blocked_payments_resolved': [resultado_final['blocked_payments_resolved'].sum()],
        'amount_resolved': [resultado_final['amount_resolved'].sum()],
        'blocked_users_resolved': [total_blocked_users_resolved],
        'new_users_blocked': [total_new_users_blocked],
        'new_users_blocked_resolved': [total_new_users_blocked_resolved],
        'total_new_users': [total_total_new_users],
        'percentage_new_users_blocked (%)': [
            0 if total_total_new_users_ajustado == 0
            else (total_new_users_blocked / total_total_new_users_ajustado) * 100
        ]
    })

    # Concatenar la fila de totales al final del DataFrame
    resultado_final = pd.concat([resultado_final, totales], ignore_index=True)

    # # Eliminar columnas innecesarias de la tabla final
    # resultado_final = resultado_final.drop(columns=['total_new_users'])

    # Guardar el resultado en un archivo Excel
    columns_to_plot = ['real_block_payments', 'real_amount_blocked', 'blocked_users', 'blocked_payments_resolved', 'amount_resolved', 'blocked_users_resolved', 'new_users_blocked', 'new_users_blocked_resolved', 'total_new_users', 'percentage_new_users_blocked (%)']  # Columnas para graficar
    colors = ['#0000FF', '#008000', '#FF0000', '#6F4E37', '#FFA500', '#800080', '#00FFFF', '#FF00FF', '#FFC0CB', '#ff8b00']  # Colores para los gráficos
    grafico_positions = ['M2', 'M24', 'M46', 'M68', 'M90', 'M112', 'M134', 'M156', 'M178', 'M200']  # Posiciones de los gráficos en el Excel

    save_dataframe_to_excel(carpeta_salida, nombre_salida, resultado_final, 'General', columns_to_plot, colors, grafico_positions)

    # Mostrar la tabla final
    print(resultado_final)