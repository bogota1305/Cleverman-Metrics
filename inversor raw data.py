import pandas as pd
from pathlib import Path

from modules.database_queries import execute_query

# =========================
# CONFIG
# =========================
OUTPUT_FILE = "cohort_metrics_2021_2025.xlsx"

COL_USER = "user_id"
COL_COHORT = "cohort"
COL_ORDERS = "total_orders_up_today"
COL_REVENUE = "total_paid_up_today"
COL_REPURCHASE_REVENUE = "repurchase_total_paid_up_today"
COL_ACTIVE_SUB = "has_active_subscription"
COL_LAST_ORDER_DATE = "last_order_date"
COL_EXPERIENCE = "experience_with_color"
COL_LAST_CANCELATION_DATE = "last_cancellation_date"

MONTH_START = "2021-01-01"
MONTH_END = "2025-12-01"  

# =========================
# QUERY
# =========================
query = f"""
    SELECT c.id AS user_id,DATE(fs.first_sub_created_at) AS cohort,IFNULL(o.total_orders_up_today,0) AS total_orders_up_today,IFNULL(o.total_paid_up_today,0) AS total_paid_up_today,IFNULL(o.total_paid_up_today,0)-IFNULL(fs.first_sub_total_paid,0) AS repurchase_total_paid_up_today,(IFNULL(s.active_subs_count,0)>0) AS has_active_subscription,DATE(lo.last_order_created_at) AS last_order_date,DATE(lc.last_cancellation_created_at) AS last_cancellation_date,CASE WHEN exp.exp_code IN ('118','112') THEN 'Currently Dyed' WHEN exp.exp_code IN ('119','116') THEN 'I''ve colored' WHEN exp.exp_code IN ('120','114') THEN 'Never colored' ELSE NULL END AS experience_with_color
    FROM prod_sales_and_subscriptions.customers c
    JOIN (SELECT t.customer_id AS user_id,t.first_sub_created_at,t.id,MAX(CASE WHEN t.created_at=t.first_sub_created_at THEN t.total ELSE 0 END) AS first_sub_total_paid FROM (SELECT fo.customer_id,fo.created_at,fo.total,fo.id,MIN(fo.created_at) OVER (PARTITION BY fo.customer_id) AS first_sub_created_at FROM bi.fact_orders fo WHERE fo.order_plan IN ('SUBSCRIPTION','MIXED') AND fo.status NOT IN ('CANCELLED','PAYMENT_ERROR')) t GROUP BY t.customer_id,t.first_sub_created_at) AS fs ON fs.user_id=c.id
    LEFT JOIN prod_sales_and_subscriptions.sales_order_items s2 ON s2.salesOrderId=fs.id AND s2.itemId LIKE "%0001004170%"
    LEFT JOIN prod_sales_and_subscriptions.sales_orders s3 ON s3.id=fs.id
    LEFT JOIN (SELECT fo.customer_id AS user_id,COUNT(*) AS total_orders_up_today,SUM(fo.total) AS total_paid_up_today FROM bi.fact_orders fo WHERE fo.status NOT IN ('CANCELLED','PAYMENT_ERROR') GROUP BY fo.customer_id) AS o ON o.user_id=c.id
    LEFT JOIN (SELECT fo.customer_id AS user_id,MAX(fo.created_at) AS last_order_created_at FROM bi.fact_orders fo WHERE fo.status NOT IN ('CANCELLED','PAYMENT_ERROR') AND fo.order_plan='OTO' GROUP BY fo.customer_id) AS lo ON lo.user_id=c.id
    LEFT JOIN (SELECT s.customerId AS user_id,COUNT(*) AS active_subs_count FROM prod_sales_and_subscriptions.subscriptions s WHERE s.status<>'CANCELLED' GROUP BY s.customerId) AS s ON s.user_id=c.id
    LEFT JOIN (SELECT s.customerId AS user_id,MAX(fc.createdAt) AS last_cancellation_created_at FROM prod_sales_and_subscriptions.subscriptions s JOIN bi.fact_cancellations fc ON fc.subscriptionId=s.id GROUP BY s.customerId) AS lc ON lc.user_id=c.id
    LEFT JOIN (SELECT d.user_id,COALESCE(d.v40_array,d.v40_obj,d.v36_array,d.v36_obj) AS exp_code FROM (SELECT fs2.user_id,(SELECT jt.val FROM JSON_TABLE(CAST(IFNULL(s2x.additionalFields->>"$.diagnostic",s3x.additionalFields->>"$.diagnostic") AS JSON),'$.values[*]' COLUMNS(var INT PATH '$.variable',val VARCHAR(255) PATH '$.value')) jt WHERE jt.var=40 LIMIT 1) AS v40_array,(SELECT jt.val FROM JSON_TABLE(CAST(IFNULL(s2x.additionalFields->>"$.diagnostic",s3x.additionalFields->>"$.diagnostic") AS JSON),'$.values.*' COLUMNS(var INT PATH '$.variable',val VARCHAR(255) PATH '$.value')) jt WHERE jt.var=40 LIMIT 1) AS v40_obj,(SELECT jt.val FROM JSON_TABLE(CAST(IFNULL(s2x.additionalFields->>"$.diagnostic",s3x.additionalFields->>"$.diagnostic") AS JSON),'$.values[*]' COLUMNS(var INT PATH '$.variable',val VARCHAR(255) PATH '$.value')) jt WHERE jt.var=36 LIMIT 1) AS v36_array,(SELECT jt.val FROM JSON_TABLE(CAST(IFNULL(s2x.additionalFields->>"$.diagnostic",s3x.additionalFields->>"$.diagnostic") AS JSON),'$.values.*' COLUMNS(var INT PATH '$.variable',val VARCHAR(255) PATH '$.value')) jt WHERE jt.var=36 LIMIT 1) AS v36_obj FROM (SELECT t.customer_id AS user_id,t.id FROM (SELECT fo.customer_id,fo.id,MIN(fo.created_at) OVER (PARTITION BY fo.customer_id) AS first_sub_created_at FROM bi.fact_orders fo WHERE fo.order_plan IN ('SUBSCRIPTION','MIXED') AND fo.status NOT IN ('CANCELLED','PAYMENT_ERROR')) t GROUP BY t.customer_id) fs2 LEFT JOIN prod_sales_and_subscriptions.sales_order_items s2x ON s2x.salesOrderId=fs2.id AND s2x.itemId LIKE "%0001004170%" LEFT JOIN prod_sales_and_subscriptions.sales_orders s3x ON s3x.id=fs2.id) d) AS exp ON exp.user_id=c.id
    WHERE DATE(fs.first_sub_created_at) BETWEEN '2021-01-01' AND '2026-01-01'
    ORDER BY cohort,user_id;
"""

df = execute_query(query)
df = df.drop_duplicates(subset=[COL_USER])

# =========================
# VALIDATE + NORMALIZE
# =========================
required = [
    COL_USER,
    COL_COHORT,
    COL_ORDERS,
    COL_REVENUE,
    COL_REPURCHASE_REVENUE,
    COL_ACTIVE_SUB,
    COL_LAST_ORDER_DATE,
    COL_EXPERIENCE,
    COL_LAST_CANCELATION_DATE,
]
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError(f"Faltan columnas requeridas: {missing}\nColumnas encontradas: {list(df.columns)}")

# cohort -> datetime -> cohort_month
df[COL_COHORT] = pd.to_datetime(df[COL_COHORT], errors="coerce")
if df[COL_COHORT].isna().any():
    bad = df.loc[df[COL_COHORT].isna(), [COL_USER, COL_COHORT]].head(10)
    raise ValueError(f"Hay cohort inválidos (no parseables). Ejemplos:\n{bad}")

df["cohort_month"] = df[COL_COHORT].dt.to_period("M").dt.to_timestamp()

df[COL_ORDERS] = pd.to_numeric(df[COL_ORDERS], errors="coerce").fillna(0).astype(int)
df[COL_REVENUE] = pd.to_numeric(df[COL_REVENUE], errors="coerce").fillna(0.0)
df[COL_REPURCHASE_REVENUE] = pd.to_numeric(df[COL_REPURCHASE_REVENUE], errors="coerce").fillna(0.0)
df[COL_ACTIVE_SUB] = pd.to_numeric(df[COL_ACTIVE_SUB], errors="coerce").fillna(0).astype(int)

# last_order_date -> datetime (puede venir vacío; lo permitimos)
df[COL_LAST_ORDER_DATE] = pd.to_datetime(df[COL_LAST_ORDER_DATE], errors="coerce")
df[COL_LAST_CANCELATION_DATE] = pd.to_datetime(df[COL_LAST_CANCELATION_DATE], errors="coerce")

# =========================
# ACTIVE USER FLAG
# Active = subs activa OR compra (OTO) en los últimos 12 meses
# Si last_order_date es NaT, solo cuenta por has_active_subscription
# =========================
as_of_date = df[COL_LAST_ORDER_DATE].max()
# Si TODOS vienen vacíos, max() será NaT -> en ese caso no aplicamos ventana de 12 meses
if pd.isna(as_of_date):
    df["is_active_user"] = (df[COL_ACTIVE_SUB] == 1).astype(int)
else:
    cutoff_date = as_of_date - pd.Timedelta(days=365)
    cutoff_cancelation_date = as_of_date - pd.Timedelta(days=90)

    df["is_active_user"] = (
        (df[COL_ACTIVE_SUB] == 1) |
        (
            df[COL_LAST_ORDER_DATE].notna() &
            (df[COL_LAST_ORDER_DATE] > df[COL_COHORT]) &
            (df[COL_LAST_ORDER_DATE] >= cutoff_date)
        ) |
        (
            df[COL_LAST_CANCELATION_DATE].notna() &
            (df[COL_LAST_CANCELATION_DATE] >= cutoff_cancelation_date)
        )
    ).astype(int)

# =========================
# FILTER TO MONTH RANGE (columns in output)
# =========================
month_index = pd.date_range(MONTH_START, MONTH_END, freq="MS")
df = df[df["cohort_month"].isin(month_index)].copy()

# =========================
# BUILD OUTPUT TABLE (reusable)
# =========================
def build_out_table(df_in: pd.DataFrame, month_index: pd.DatetimeIndex) -> pd.DataFrame:
    df_local = df_in.copy()

    # Repurchases (orders) = total_orders - 1 (clipped at 0)
    df_local["repurchases_up_today_orders"] = (df_local[COL_ORDERS] - 1).clip(lower=0)

    g = df_local.groupby("cohort_month", dropna=False)

    metrics = pd.DataFrame(
    {
        "New users (or subscribers)": g[COL_USER].nunique(),
        "Orders up today": g[COL_ORDERS].sum(),
        "Repurchases up today - orders": g["repurchases_up_today_orders"].sum(),
        "Orders up today - revenue": g[COL_REVENUE].sum(),
        "Repurchases up today - revenue": g[COL_REPURCHASE_REVENUE].sum(),
        "Users who repurchase": g.apply(lambda x: (x[COL_ORDERS] >= 2).sum()),
        "Active users (up today)": g["is_active_user"].sum(),
        "Repurchases up today - orders (active users only)": g.apply(
            lambda x: x.loc[x["is_active_user"] == 1, "repurchases_up_today_orders"].sum()
        ),
    }
)

    metrics = metrics.reindex(month_index, fill_value=0)

    out = metrics.T
    out.columns = [d.strftime("%Y-%m") for d in out.columns]

    # Orden exacto solicitado
    desired_order = [
        "New users (or subscribers)",
        "Orders up today",
        "Orders up today - revenue",
        "Users who repurchase",
        "Repurchases up today - orders",
        "Repurchases up today - revenue",
        "Active users (up today)",
        "Repurchases up today - orders (active users only)"
    ]
    out = out.reindex(desired_order)

    # Cast int rows
    int_rows = [
        "New users (or subscribers)",
        "Orders up today",
        "Repurchases up today - orders",
        "Users who repurchase",
        "Active users (up today)",
        "Repurchases up today - orders (active users only)"
    ]
    for r in int_rows:
        if r in out.index:
            out.loc[r] = out.loc[r].astype(int)

    return out

# =========================
# SPLITS
# =========================
out_total = build_out_table(df, month_index)

# Never colored
df_never = df[df[COL_EXPERIENCE] == "Never colored"].copy()
out_never = build_out_table(df_never, month_index)

# I've colored (nota: en pandas es "I've colored", no "I''ve colored")
df_used_before = df[df[COL_EXPERIENCE] == "I've colored"].copy()
out_used_before = build_out_table(df_used_before, month_index)

# Currently Dyed
df_currently_dyed = df[df[COL_EXPERIENCE] == "Currently Dyed"].copy()
out_currently_dyed = build_out_table(df_currently_dyed, month_index)

# =========================
# EXPORT TO EXCEL (4 sheets)
# =========================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    out_total.to_excel(writer, sheet_name="Cohort Metrics", index=True)
    out_never.to_excel(writer, sheet_name="did not use hair color before", index=True)
    out_used_before.to_excel(writer, sheet_name="used hair color before", index=True)
    out_currently_dyed.to_excel(writer, sheet_name="use hair color currently", index=True)

    for sheet_name in [
        "Cohort Metrics",
        "did not use hair color before",
        "used hair color before",
        "use hair color currently",
    ]:
        ws = writer.book[sheet_name]
        ws.freeze_panes = "B2"

        for col_cells in ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells:
                val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 35)

print(f"OK -> Generado: {Path(OUTPUT_FILE).resolve()}")