import pandas as pd
from datetime import datetime, timezone
import mysql.connector
import os
from dotenv import load_dotenv

load_dotenv()


# =========================
# CONFIG
# =========================

REVIEWS_CSV = "verified_reviews_4_5_with_pros_cons.csv"

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": "dev_legacy"
}

TABLE_NAME = "review"
VISIBLE = 0
COMMIT_EVERY = 50


# =========================
# DB
# =========================

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)


# =========================
# HELPERS
# =========================

def iso_to_mysql_datetime(value):
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None

    try:
        if s.endswith("Z"):
            dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
        else:
            dt = datetime.fromisoformat(s)
            if dt.tzinfo:
                dt = dt.astimezone(timezone.utc).replace(tzinfo=None)

        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def normalize_legacy_array(value):
    if pd.isna(value):
        return "a:0:{}"
    v = str(value).strip()
    return v if v else "a:0:{}"


def nan_to_none(value):
    if pd.isna(value):
        return None
    v = str(value).strip()
    return v if v else None


def get_typo_id_from_sku(sku_raw) -> int:
    """
    SOLO SKU -> typo_id (sin name).
    Interpretación basada en la tabla que pasaste.
    """
    sku = str(sku_raw).strip().upper()
    if not sku or sku == "NAN":
        return 1

    # 7 - ALL IN ONE KIT
    # SKUs: 10412023,24,25,28,29,30,47,53,54,55,56,57...
    if sku.startswith("104120"):
        return 7

    # 10 - ENERGIZING FACE & BEARD SCRUB
    # SKUs: 10107001, 10107002, 10107003
    if sku.startswith("101070"):
        return 10

    # 12 - DAMAGED CONDITIONER (bucket para conditioner + hair treatment + hair care)
    # SS - hair conditioner: 10102008-*
    # BG - hair conditioner: 10102004-*
    # Hair Treatment: 10206001..10206009 (incluye 10206002/03 repetidos en tu tabla)
    # Hair Care: 10108001, 10412049, 10412050
    if sku.startswith("10102008-") or sku.startswith("10102004-"):
        return 12
    if sku.startswith("102060"):
        return 12
    if sku in {"10108001", "10412049", "10412050"}:
        return 12

    # 8 - GROOMING TOOLS (bucket para tools/accesorios/packaging)
    # gloves: 10102001-*, 10102005-*
    # brush/tray/spatula: 10102003-*, 10211002
    # beard tools: 10102007-*
    # grooming tools: 101021101, 101021102, 101021103
    # packaging: shipper/sleeves/boxes/brochure template
    if sku.startswith("10102001-") or sku.startswith("10102005-"):
        return 8
    if sku.startswith("10102003-") or sku == "10211002":
        return 8
    if sku.startswith("10102007-"):
        return 8
    if sku in {"101021101", "101021102", "101021103"}:
        return 8
    if sku in {"10101001", "10101002", "10101006", "10101007", "10101008", "10101010", "10113002", "10102006", "10203013"}:
        return 8

    # 2 - TOUCH UP KIT (bucket para Refill Kit)
    # 10102009..10102042
    if sku.isdigit():
        n = int(sku)
        if 10102009 <= n <= 10102042:
            return 2

    # 1 - BEARD KIT (bucket para beard care + componentes core + preassembled + gifts/others)
    # Beard Care: 10207005, 10207006, 101021104
    if sku in {"10207005", "10207006", "101021104"}:
        return 1

    # Components: colorant/developer/standalone kit components/preassembled/non-cleverman/gifts
    # (No se puede deducir 3/4/5/6 solo por SKU)
    if sku.startswith("10204") or sku.startswith("10205"):   # colorant / developer
        return 1
    if sku in {"10207001", "10207002", "10207003", "10208002"}:  # standalone components
        return 1
    if sku.startswith("105200"):  # preassembled
        return 1
    if sku.startswith("10316") or sku.startswith("10315") or sku.startswith("1020700") or sku.startswith("1020600"):
        return 1
    if sku == "20208001":  # Non-Cleverman Products
        return 1

    # fallback seguro
    return 1


# =========================
# MAIN
# =========================

def main():
    reviews_df = pd.read_csv(REVIEWS_CSV)

    # Normaliza columnas reviews a lower
    reviews_df = reviews_df.rename(columns={c: c.strip().lower() for c in reviews_df.columns})

    if "sku" not in reviews_df.columns:
        raise Exception(f"El CSV de reviews debe traer columna 'sku'. Encontré: {list(reviews_df.columns)}")

    conn = get_connection()
    cursor = conn.cursor()

    # Carga typo_ids válidos para evitar FK
    cursor.execute("SELECT id FROM dev_legacy.typo")
    valid_typo_ids = {int(r[0]) for r in cursor.fetchall()}

    insert_sql = f"""
        INSERT INTO {TABLE_NAME}
        (
            typo_id, headline, comment, nickname, email,
            pros, cons, recomendation, visible,
            updated_at, overall_rating, date
        )
        VALUES
        (
            %(typo_id)s, %(headline)s, %(comment)s, %(nickname)s, %(email)s,
            %(pros)s, %(cons)s, %(recomendation)s, %(visible)s,
            %(updated_at)s, %(overall_rating)s, %(date)s
        )
    """

    inserted = 0

    try:
        for i, r in reviews_df.iterrows():
            sku = str(r.get("sku", "")).strip().upper()
            typo_id = get_typo_id_from_sku(sku)

            if typo_id not in valid_typo_ids:
                print(f"[WARN] typo_id={typo_id} no existe en dev_legacy.typo. Fallback a 1. sku={sku} row={i}")
                typo_id = 1

            data = {
                "typo_id": int(typo_id),
                "headline": nan_to_none(r.get("headline")),
                "comment": nan_to_none(r.get("comment")),
                "nickname": nan_to_none(r.get("nickname")),
                "email": nan_to_none(r.get("email")),
                "pros": normalize_legacy_array(r.get("pros")),
                "cons": normalize_legacy_array(r.get("cons")),
                "recomendation": int(r["recomendation"]) if not pd.isna(r.get("recomendation")) else None,
                "visible": VISIBLE,
                "updated_at": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                "overall_rating": int(r["overall_rating"]) if not pd.isna(r.get("overall_rating")) else None,
                "date": iso_to_mysql_datetime(r.get("date")),
            }

            try:
                cursor.execute(insert_sql, data)
                inserted += 1
            except mysql.connector.Error as ex:
                print(f"[ERROR] Fila {i} falló. sku={sku} typo_id={typo_id}. Error={ex}")
                raise

            if inserted % COMMIT_EVERY == 0:
                conn.commit()
                print(f"{inserted} reviews insertadas...")

        conn.commit()
        print(f"✔ Total reviews insertadas: {inserted}")

    except Exception as e:
        conn.rollback()
        raise e
    finally:
        cursor.close()
        conn.close()


if __name__ == "__main__":
    main()
