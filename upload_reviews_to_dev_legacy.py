import pandas as pd
from datetime import datetime, timezone
import mysql.connector
import os
from dotenv import load_dotenv

load_dotenv()

# =========================
# CONFIG
# =========================

REVIEWS_CSV = "verified_reviews_processed.csv"

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": "prod_legacy"
}

TABLE_NAME = "review"
VISIBLE = 0
COMMIT_EVERY = 50

# =========================
# SKU -> PICTURE
# =========================
SKU_TO_PICTURE = {
    "10412029": "lightestblond.jpg",
    "10412028": "mediumBlond.jpg",
    "10412030": "auburn.jpg",
    "10412057": "lightBrown.jpg",
    "10412056": "mediumBrown.jpg",
    "10412025": "darkBrown.jpg",
    "10412024": "black.jpg",
    "10412023": "jetBlack.jpg",
    "10412047": "darkblond.jpg",
    "10412053": "blackAfro.jpg",
    "10412054": "Jetbalckafro.jpg",
    "10412055": "darkBrownAfro.jpg",
    "10207005": "sensitiveScrub.jpg",
    "101021103": "Scissors.jpg",
    "101021110": "Clipper.jpg",
    "10412049": "CleansingShampooforNormal.jpg",
    "10412050": "CleansingConditionerforDamagedHair.jpg",
    "10206003": "CleansingConditionerforDamagedHair.jpg",

    "10511117": "lightestblond.jpg",
    "10511116": "mediumBlond.jpg",
    "10511118": "auburn.jpg",
    "10511115": "lightBrown.jpg",
    "10511114": "mediumBrown.jpg",
    "10511113": "darkBrown.jpg",
    "10511112": "black.jpg",
    "10511111": "jetBlack.jpg",
    "10511119": "darkblond.jpg",
    "10511120": "blackAfro.jpg",
    "10511121": "Jetbalckafro.jpg",
    "10511122": "darkBrownAfro.jpg",

    "10412105": "blackAfro.jpg",
    "10412101": "mediumBlond.jpg",
    "10412100": "lightBrown.jpg",
    "10412099": "mediumBrown.jpg",
    "10412098": "darkBrown.jpg",
    "10412097": "black.jpg",
    
    "10412086": "blackAfro2x.jpg",
    "10412085": "lightestblond2x.jpg",
    "10412083": "auburn2x.jpg",
    "10412082": "darkblond2x.jpg",
    "10412080": "mediumBrown2x.jpg",
    "10412079": "darkBrown2x.jpg",
    "10412078": "jetBlack2x.jpg",
    "10412077": "black2x.jpg",
}

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


def get_picture_from_sku(sku_raw):

    sku = str(sku_raw).strip().upper()
    if not sku or sku == "NAN":
        return None
    return SKU_TO_PICTURE.get(sku)


def normalize_nickname(value):
    if pd.isna(value):
        return None

    s = str(value).strip()
    if not s:
        return None

    return s.split(" ")[0]


def get_typo_id_from_sku(sku_raw) -> int:

    sku = str(sku_raw).strip().upper()
    if not sku or sku == "NAN":
        return 1

    # 7 - ALL IN ONE KIT
    if sku.startswith("104120") or sku.startswith("105111"):
        return 7

    # 10 - ENERGIZING FACE & BEARD SCRUB
    if sku.startswith("101070"):
        return 10

    # 12 - DAMAGED CONDITIONER (bucket para conditioner + hair treatment + hair care)
    if sku.startswith("10102008-") or sku.startswith("10102004-"):
        return 12
    if sku.startswith("102060"):
        return 12
    if sku in {"10108001", "10412049", "10412050", "10206003"}:
        return 12

    # 8 - GROOMING TOOLS (bucket para tools/accesorios/packaging)
    if sku.startswith("10102001-") or sku.startswith("10102005-"):
        return 8
    if sku.startswith("10102003-") or sku == "10211002":
        return 8
    if sku.startswith("10102007-"):
        return 8
    if sku in {"101021101", "101021102", "101021103"}:
        return 8
    if sku in {"10101001", "10101002", "10101006", "10101007", "10101008", "10101010",
               "10113002", "10102006", "10203013"}:
        return 8

    # 2 - TOUCH UP KIT (bucket para Refill Kit)
    if sku.isdigit():
        n = int(sku)
        if 10102009 <= n <= 10102042:
            return 2

    # 1 - BEARD KIT (bucket para beard care + componentes core + preassembled + gifts/others)
    if sku in {"10207005", "10207006", "101021104"}:
        return 1

    # Components: colorant/developer/standalone kit components/preassembled/non-cleverman/gifts
    if sku.startswith("10204") or sku.startswith("10205"):
        return 1
    if sku in {"10207001", "10207002", "10207003", "10208002"}:
        return 1
    if sku.startswith("105200"):
        return 1
    if sku.startswith("10316") or sku.startswith("10315") or sku.startswith("1020700") or sku.startswith("1020600"):
        return 1
    if sku == "20208001":
        return 1

    return 1


# =========================
# MAIN
# =========================
def main():
    reviews_df = pd.read_csv(REVIEWS_CSV)

    reviews_df = reviews_df.rename(columns={c: c.strip().lower() for c in reviews_df.columns})

    if "sku" not in reviews_df.columns:
        raise Exception(f"El CSV de reviews debe traer columna 'sku'. Encontré: {list(reviews_df.columns)}")

    conn = get_connection()
    cursor = conn.cursor()

    insert_sql = f"""
        INSERT INTO {TABLE_NAME}
        (
            typo_id, headline, comment, nickname, email,
            picture, pros, cons, recomendation, visible,
            updated_at, overall_rating, date
        )
        VALUES
        (
            %(typo_id)s, %(headline)s, %(comment)s, %(nickname)s, %(email)s,
            %(picture)s, %(pros)s, %(cons)s, %(recomendation)s, %(visible)s,
            %(updated_at)s, %(overall_rating)s, %(date)s
        )
    """

    inserted = 0

    try:
        for i, r in reviews_df.iterrows():
            sku = str(r.get("sku", "")).strip().upper()

            typo_id = get_typo_id_from_sku(sku)
            picture = get_picture_from_sku(sku)

            data = {
                "typo_id": int(typo_id),
                "headline": nan_to_none(r.get("headline")),
                "comment": nan_to_none(r.get("comment")),
                "nickname": normalize_nickname(r.get("nickname")),
                "email": nan_to_none(r.get("email")),
                "picture": picture,
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
                print(f"[ERROR] Fila {i} falló. sku={sku} typo_id={typo_id} picture={picture}. Error={ex}")
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
