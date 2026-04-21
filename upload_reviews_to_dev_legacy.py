import pandas as pd
from datetime import datetime, timezone
import mysql.connector
import os
from dotenv import load_dotenv
import random
import string

load_dotenv()

# =========================
# CONFIG
# =========================

REVIEWS_CSV = "verified_reviews_final_latest_v2.csv"

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": "prod_ecommerce"
}

TABLE_NAME = "review"
VISIBLE = 0
COMMIT_EVERY = 50
CREATED_BY = "amazon"

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
    "10412102": "lightestblond.jpg",
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


def generate_review_id():
    chars = string.ascii_uppercase + string.digits
    return "REV" + ''.join(random.choices(chars, k=29))


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


def get_productReviewTypeId_from_sku(sku_raw) -> int:
    sku = str(sku_raw).strip().upper()
    if not sku or sku == "NAN":
        return 1

    # 7 - ALL IN ONE KIT
    if sku.startswith("104120") or sku.startswith("105111"):
        return 7

    # 10 - ENERGIZING FACE & BEARD SCRUB
    if sku.startswith("101070"):
        return 10

    # 12 - DAMAGED CONDITIONER
    if sku.startswith("10102008-") or sku.startswith("10102004-"):
        return 12
    if sku.startswith("102060"):
        return 12
    if sku in {"10108001", "10412049", "10412050", "10206003"}:
        return 12

    # 8 - GROOMING TOOLS
    if sku.startswith("10102001-") or sku.startswith("10102005-"):
        return 8
    if sku.startswith("10102003-") or sku == "10211002":
        return 8
    if sku.startswith("10102007-"):
        return 8
    if sku in {"101021101", "101021102", "101021103"}:
        return 8
    if sku in {
        "10101001", "10101002", "10101006", "10101007", "10101008",
        "10101010", "10113002", "10102006", "10203013"
    }:
        return 8

    # 2 - TOUCH UP KIT
    if sku.isdigit():
        n = int(sku)
        if 10102009 <= n <= 10102042:
            return 2

    # 1 - BEARD KIT
    if sku in {"10207005", "10207006", "101021104"}:
        return 1

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


def build_product_review_type_id(type_id: int) -> str:
    if type_id > 9:
        return f"PRT000000000000000000000000000{type_id}"
    return f"PRT0000000000000000000000000000{type_id}"


def parse_int_or_none(value):
    if pd.isna(value):
        return None

    s = str(value).strip()
    if not s:
        return None

    try:
        return int(float(s))
    except Exception:
        return None


# =========================
# MAIN
# =========================
def main():
    reviews_df = pd.read_csv(REVIEWS_CSV)

    # Normalizar nombres de columnas
    reviews_df = reviews_df.rename(columns={c: c.strip().lower() for c in reviews_df.columns})

    print("Columnas detectadas en el CSV:")
    print(reviews_df.columns.tolist())

    if "sku" not in reviews_df.columns:
        raise Exception(f"El CSV de reviews debe traer columna 'sku'. Encontré: {list(reviews_df.columns)}")

    if "overallrating" not in reviews_df.columns:
        raise Exception(f"El CSV no trae la columna obligatoria 'overallrating'. Encontré: {list(reviews_df.columns)}")

    if "date" not in reviews_df.columns:
        raise Exception(f"El CSV no trae la columna obligatoria 'date'. Encontré: {list(reviews_df.columns)}")

    conn = get_connection()
    cursor = conn.cursor()

    insert_sql = f"""
        INSERT INTO {TABLE_NAME}
        (
            id,
            createdBy,
            productReviewTypeId, headline, comment, nickname, email,
            additionalFields, pros, cons, recommendation, visible,
            updatedAt, overallRating, reviewDate, createdAt
        )
        VALUES
        (
            %(id)s,
            %(createdBy)s,
            %(productReviewTypeId)s, %(headline)s, %(comment)s, %(nickname)s, %(email)s,
            %(additionalFields)s, %(pros)s, %(cons)s, %(recommendation)s, %(visible)s,
            %(updatedAt)s, %(overallRating)s, %(reviewDate)s, %(createdAt)s
        )
    """

    inserted = 0

    try:
        for i, r in reviews_df.iterrows():
            sku = str(r.get("sku", "")).strip().upper()

            type_id = get_productReviewTypeId_from_sku(sku)
            product_review_type_id = build_product_review_type_id(type_id)

            picture = get_picture_from_sku(sku)
            additional_fields = (
                f'{{"mainImage": "https://cdn.becleverman.com/uploads/images/reviews/{picture}"}}'
                if picture
                else None
            )

            # =========================
            # DATE (FUENTE ÚNICA)
            # =========================
            date_raw = r.get("date")
            parsed_date = iso_to_mysql_datetime(date_raw)

            if parsed_date is None:
                print(f"[ERROR] Fila {i} con date inválido. email={r.get('email')} valor={date_raw}")
                raise Exception(f"Fila {i} sin fecha válida")

            review_date = parsed_date
            created_at = parsed_date

            # =========================
            # OTROS CAMPOS
            # =========================
            overall_rating = parse_int_or_none(r.get("overallrating"))
            recommendation = parse_int_or_none(r.get("recommendation"))

            if overall_rating is None:
                print(f"[ERROR] Fila {i} sin overallrating válido. email={r.get('email')} sku={sku}")
                raise Exception(f"Fila {i} sin overallrating")

            data = {
                "id": generate_review_id(),
                "createdBy": CREATED_BY,
                "productReviewTypeId": product_review_type_id,
                "headline": nan_to_none(r.get("headline")),
                "comment": nan_to_none(r.get("comment")),
                "nickname": normalize_nickname(r.get("nickname")),
                "email": nan_to_none(r.get("email")),
                "additionalFields": additional_fields,
                "pros": normalize_legacy_array(r.get("pros")),
                "cons": normalize_legacy_array(r.get("cons")),
                "recommendation": recommendation,
                "visible": VISIBLE,
                "updatedAt": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S"),
                "overallRating": overall_rating,
                "reviewDate": review_date,
                "createdAt": created_at,
            }

            try:
                cursor.execute(insert_sql, data)
                inserted += 1
            except mysql.connector.Error as ex:
                print(
                    f"[ERROR] Fila {i} falló. "
                    f"sku={sku} "
                    f"id={data['id']} "
                    f"createdBy={data['createdBy']} "
                    f"productReviewTypeId={product_review_type_id} "
                    f"picture={picture} "
                    f"overallRating={overall_rating} "
                    f"recommendation={recommendation} "
                    f"reviewDate={review_date} "
                    f"createdAt={created_at} "
                    f"Error={ex}"
                )
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