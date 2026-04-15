# Cleverman Metrics – Business Reporting Suite

## Overview

This project provides a collection of Python scripts designed to extract, process, and visualize business metrics from Cleverman's database. The primary goals are:

1. Generate comprehensive **Monthly Reports** in Excel (orders, sales, renewals, payment errors, funnels, and more).
2. Generate the **Full Control (FC) Report** that tracks subscription renewal behavior.
3. Read and upload **customer reviews** from the SUVAE platform to the production database.
4. Analyze **repurchase behavior and subscription cancellations** on a quarterly basis.

---

## Table of Contents

1. [Prerequisites & Installation](#1-prerequisites--installation)
2. [Configuration](#2-configuration)
3. [Pipeline Frequency & Schedule](#3-pipeline-frequency--schedule)
4. [Monthly Report – `main.py`](#4-monthly-report--mainpy)
5. [Full Control Report – `fcReport.py`](#5-full-control-report--fcreportpy)
6. [Reviews Pipeline](#6-reviews-pipeline)
   - [Step 1 – Read Reviews – `read_reviews.py`](#step-1--read-reviews--read_reviewspy)
   - [Step 2 – AI Enrichment (pros & cons)](#step-2--ai-enrichment-pros--cons)
   - [Step 3 – Upload Reviews – `upload_reviews_to_dev_legacy.py`](#step-3--upload-reviews--upload_reviews_to_dev_legacypy)
7. [Repurchase & Cancellations Pipeline](#7-repurchase--cancellations-pipeline)
   - [Step 1 – Repurchase Diagnostic – `repurchaseFirstOrderDiagnosticTotal.py`](#step-1--repurchase-diagnostic--repurchasefirstorderdiagnostictotalpy)
   - [Step 2 – Shade Cancellations – `shadeCancelations.py`](#step-2--shade-cancellations--shadecancelationspy)
   - [Step 3 – Merge Report – `analisis_repurchase_cancelaciones.py`](#step-3--merge-report--analisis_repurchase_cancelacionespy)
8. [Supporting Modules](#8-supporting-modules)
9. [Cloud Upload](#9-cloud-upload)
10. [General Metrics Report](#10-general-metrics-report)
11. [Project Structure](#11-project-structure)

---

## 1. Prerequisites & Installation

### Python Version

- Python **3.10+** (type hints such as `dict[str, ...]` are used in several files)

### Required Libraries

Install all dependencies with a single command:

```bash
pip install pandas mysql-connector-python matplotlib openpyxl tkcalendar xlsxwriter \
            google-api-python-client google-auth-httplib2 google-auth-oauthlib \
            dropbox python-dotenv numpy
```

> **Note:** `tkinter` is part of the Python standard library on most systems. On some Linux distributions you may need to install it separately (`sudo apt-get install python3-tk`).

### Summary of Dependencies by Script

| Library | Used by |
|---|---|
| `pandas` | All scripts |
| `mysql-connector-python` | `main.py`, `fcReport.py`, `upload_reviews_to_dev_legacy.py`, and all DB-querying scripts |
| `openpyxl` | `fcReport.py`, `read_reviews.py`, `excel_creator.py`, `analisis_repurchase_cancelaciones.py` |
| `matplotlib` | `excel_creator.py` and chart-generating scripts |
| `tkinter` + `tkcalendar` | `date_selector.py`, `selectFiles.py` |
| `xlsxwriter` | `ga4Funnels.py` |
| `google-api-python-client`, `google-auth-*` | `uploadCloud.py` |
| `dropbox` | `uploadCloud.py` |
| `python-dotenv` | `upload_reviews_to_dev_legacy.py` |
| `numpy` | `shadeCancelations.py` |

---

## 2. Configuration

### Database Connection

Database credentials are stored in a **`.env`** file at the root of the repository:

```
DB_HOST = <aurora-mysql-host>
DB_USER = <db-username>
DB_PASSWORD = <db-password>
```

The scripts load these values automatically via `python-dotenv`. Make sure the `.env` file exists before running any script that queries the database.

For the main reporting scripts (`main.py`, etc.), the database connection is handled inside `modules/database_queries.py`. If you need to change the connection details there instead, open `modules/database_queries.py` and update the `host`, `user`, and `password` variables.

> ⚠️ **Never commit the `.env` file or `credentials.json` to version control.** They are already listed in `.gitignore`.

---

## 3. Pipeline Frequency & Schedule

| Pipeline | Frequency | Notes |
|---|---|---|
| **Monthly Report** (`main.py`) | Once a month, at the **beginning** of the month | Covers the previous month's data |
| **Full Control Report** (`fcReport.py`) | Once a month, at the **beginning** of the month | Run alongside the Monthly Report |
| **Reviews Pipeline** | Every **2 weeks** | Run `read_reviews.py` → AI enrichment → `upload_reviews_to_dev_legacy.py` |
| **Repurchase & Cancellations** | **Quarterly** — always for the **previous** quarter | See note below |

> 📅 **Quarterly pipeline timing:** This report is always generated for the quarter that has **already closed**, not the one currently in progress. For example, as soon as Q1 2026 ends (April 1, 2026), you run the report for **Q4 2025** (Oct 1 – Dec 31, 2025). You do **not** run Q1 2026 until Q2 2026 begins.

---

## 4. Monthly Report – `main.py`

`main.py` is the **main entry point** for generating the monthly business metrics report. It orchestrates multiple sub-scripts that query the database and process GA4 funnel data, then consolidates everything into a single Excel file.

### What it does

- Opens a **GUI selector** for the report type (Funnels only, Database only, or both)
- Opens a **GUI date selector** to pick the reporting period and the modules to include
- Queries the database for each selected metric and generates individual Excel files
- Processes GA4 funnel CSV files (downloaded from Google Analytics)
- Records all metrics into the shared `metricas.xlsx` tracker
- Optionally uploads generated files to **Dropbox** or **Google Drive**
- Produces a final `Monthly Report <Month>.xlsx` file

### GA4 Funnels – Important Prerequisite

Before running `main.py`, you must **manually download the funnel CSV files from Google Analytics 4 (GA4)**:

1. Go to [Google Analytics](https://analytics.google.com/analytics/web/?authuser=1#/p338732175/reports/reportinghub)
2. Navigate to the desired funnel report (e.g. *Customized Kit*, *All In One*, *Shop*, etc.)
3. Click the **Download** button and select **CSV** format

   ![GA4 Download Example](https://github.com/user-attachments/assets/5bbb15ef-a7e8-4d66-af32-aaf4b38d0e17)

4. Repeat for each funnel you want to include in the report
5. When `main.py` prompts you to select files, pick the corresponding CSV for each funnel

### Supported funnels

| Funnel Name | Description |
|---|---|
| Customized Kit - Funnel | Main kit purchase funnel |
| All In One - Funnel | All-in-one kit funnel |
| Shop - Funnel | Store funnel |
| My Account - Funnel | Account management funnel |
| Buy Again - Funnel | Repurchase funnel |
| My Subscriptions - Funnel | Subscription management |
| My Subscriptions Reactivate - Funnel | Reactivation flow |
| My Subscriptions Without Sub - Funnel | Non-subscriber flow |
| NPD mail - Funnel | NPD email flow |
| NPD account - Funnel | NPD account flow |
| Beard / Hair landing page funnels | Multiple color variants |

### Report sections (selectable in GUI)

| Option | Script called | Description |
|---|---|---|
| All orders | `orders.py` | New & existing user orders |
| Unique orders | `orders.py` | De-duplicated orders |
| Sales | `renewalsAndNoRecurrents.py` | Recurring & non-recurring sales |
| Payment errors | `payments.py` | Payment failure analysis |
| Expected renewals | `exceptedRenewals.py` | Renewal forecasting |
| Renewal frequency | `realRenewalFrecuency.py` | Frequency of renewals |
| Full control | `fullContol.py` | FC subscription metrics |
| Subscriptions | `subscriptions.py` | Subscription breakdown |
| Refill | `refill.py` | Refill orders |
| Upsize | `upsize.py` | Upsell metrics |
| How did you hear from us | `howHearFromUs.py` | Attribution survey |

### How to run

```bash
python main.py
```

A series of GUI windows will appear guiding you through:
1. Selecting the report type (funnels / database / both)
2. Selecting the GA4 funnel CSV files (if applicable)
3. Selecting the date range and report modules (if database is selected)
4. Selecting the cloud upload destination (Dropbox / Google Drive / local only)

> **Date selection tip:** Choose the end date as the day **after** the last day you want included (e.g. for Oct 1–10, set start = Oct 1 and end = Oct 11).

### Where to upload

- All generated files are uploaded to **your own Google Drive** automatically via `uploadCloud.py`.
- The **main report file** (`Monthly Report <Month>.xlsx`) must be **moved manually** to the following shared Drive folder:
  [📁 Monthly Reports Folder](https://drive.google.com/drive/folders/1wEyoOxKCOkCWl6-tzkkgoXQiw-jsSrQq?usp=drive_link)

---

## 5. Full Control Report – `fcReport.py`

`fcReport.py` generates the **Full Control (FC) tracker**, which measures the behavior of subscribers enrolled in the SMS renewal program.

### What it does

- Queries the `prod_sales_and_subscriptions` schema to get the history of subscriptions enrolled in Full Control (via `first_sms_renewal_versions` and `renewals` tables)
- Calculates enriched columns:
  - **Unique Subscription Flag** – marks the first appearance of each subscription
  - **Is Reactivation Renewal** – flags renewals that happened on the same day as enrollment
  - **Reactiv renewal 1** – identifies subscriptions that reactivated on day 1
- Builds **10 monthly aggregated tables**:

| Table | Description |
|---|---|
| `customers_joined_program` | Total new subscribers enrolled per month |
| `reactivation_renewals` | Subscribers who enrolled and bought the same day |
| `enrolled_bought_same_day_and_bought_more_than_once` | Same-day buyers who also renewed at least once more |
| `first_renewal_not_reactivation` | Subscribers whose first renewal happened on a different day |
| `no_renewals_yet` | Subscribers who enrolled but never renewed |
| `unique_subscriptions_active_processing` | Active + Processing subscriptions per cohort month |
| `unique_subscriptions_active_processing_onhold` | Active + Processing + On Hold |
| `all_renewals_after_enrollment` | Total program renewals |
| `second_renewal` | Second renewal per subscriber |
| `second_or_more_renewals` | Renewals #2 and beyond |

- Opens the Excel template **`Ecomm initiatives trackers.xlsx`** and fills in the **"Full control"** sheet by matching months from the template headers (row 2, columns B–Z) with the calculated data
- Saves the result as **`Ecomm initiatives trackers - filled.xlsx`**

### How to run

```bash
python fcReport.py
```

**Prerequisites:**
- The file `Ecomm initiatives trackers.xlsx` must exist in the same folder as `fcReport.py`
- The `.env` file must contain valid database credentials

### Template requirements

The script reads month headers from **row 2, columns B through Z** of the "Full control" sheet. Months must be formatted as dates (e.g. `2024-01-01` or any format parseable by `pandas.to_datetime`). The script writes values into the following rows:

| Row | Metric |
|---|---|
| 4 | Customers joined the program |
| 15 | Reactivation renewals |
| 16 | Enrolled + bought same day + bought again |
| 20 | First renewal (not a reactivation) |
| 21 | No renewals yet |
| 24 | Active + Processing |
| 26 | Active + Processing + On Hold |
| 40 | All renewals after enrollment |
| 42 | Second renewal |
| 43 | Second or more renewals |

### Where to upload

The output of `fcReport.py` is used to fill the **Full Control sheet** in the following shared Drive spreadsheet:

[📊 Ecomm Initiatives Trackers (Full Control sheet)](https://docs.google.com/spreadsheets/d/1t8fJrSbNgA51psvYKXwb35wcG3n8DS2oHWPJkSczaUI/edit?usp=sharing)

> **Note:** Only the **"Full control"** sheet in that spreadsheet is relevant to this report. The other sheets are managed by other processes.

---

## 6. Reviews Pipeline

The reviews pipeline consists of **three sequential steps** to import customer reviews from the SUVAE platform into the production database.

---

### Step 1 – Read Reviews – `read_reviews.py`

#### Where to get the JSON

1. Go to the **SUVAE** platform review page in your browser
2. Open the **browser developer tools** (F12) and navigate to the **Network** tab
3. Reload the page and look for a network request that returns the reviews data (usually a `GET` request that returns a JSON with an `itemsList` field)
4. Copy the response body (the full JSON)
5. Save it as **`reviews.json`** in the root of the repository

#### What the script does

- Reads `reviews.json`
- Filters only reviews that have `adminStatus` containing `"VERIFIED"` and a rating of **4 or 5 stars**
- Maps star ratings to `recommendation` values (5★ → 10, 4★ → 8)
- Extracts fields: `Product`, `SKU`, `Order ID`, `headline`, `comment`, `nickname`, `email`, `recommendation`, `overallrating`, `date`
- Generates two output files:
  - **`verified_reviews_4_5.xlsx`** – Excel workbook with the filtered reviews and a summary sheet
  - **`verified_reviews_4_5.csv`** – CSV version of the same data (used in Step 3)

#### How to run

```bash
python read_reviews.py
```

---

### Step 2 – AI Enrichment (pros & cons)

After running `read_reviews.py`, the resulting CSV/Excel file **does not yet contain the `pros` and `cons` columns** required by the database.

You must use an **AI tool** (e.g. ChatGPT, Claude, Gemini, etc.) to generate these columns.

> ⚠️ **Privacy note:** Before uploading the file to any AI tool, **remove any columns containing sensitive user data** such as email addresses, full names, or order IDs. The only field needed by the AI is the review `comment` text.

#### Valid values for `pros` and `cons`

The `pros` and `cons` columns must use **only** the following predefined values. The AI must choose from this exact list — it cannot invent new categories.

**Valid `pros` values:**
```
Natural-Looking Result
Eliminates Desired Level Of Gray
Long-Lasting
Color Matches Desired Result
Conditioner Makes My Hair Feel Great
Brush Makes Application Easy
High Quality Gloves That Fit My Hands
Customized Instructions, Easy To Understand
Auto-Delivery
Mess-Free
Works Fast
```

**Valid `cons` values:**
```
Covered Too Much Gray
Covered Too Little Gray
Fades Quickly
Color Was Lighter Than I Expected
Color Was Darker Than I Expected
Difficult To Apply
Confusing Instructions
Messy
Time-Consuming
Stains In Skin / Sink
```

> **Format:** Both `pros` and `cons` are stored as **JSON arrays** in the database (e.g. `["Natural-Looking Result", "Mess-Free"]`). The AI must output this exact format and the upload script reads it directly — no conversion needed.

#### AI Prompt

Copy and paste the following prompt directly into any AI tool along with the reviews file (with sensitive columns removed):

---

> You are a data enrichment assistant for a men's hair coloring brand. I will provide you with a CSV file containing customer reviews. Your task is to analyze the `comment` field of each review and add two new columns: `pros` and `cons`.
>
> **Instructions:**
>
> 1. Read the `comment` field of each review carefully.
> 2. Based on the content of the comment, select the most appropriate values from the **predefined lists below**. Do not invent new categories — only use values from these lists.
> 3. If a review has no negative aspects mentioned, leave `cons` as an empty list `[]`.
> 4. Output the result as a new CSV file with all original columns preserved, plus the two new columns `pros` and `cons`.
> 5. The format for `pros` and `cons` must be a **JSON array of strings**, for example:
>    `["Natural-Looking Result", "Mess-Free"]`
>
> **Valid values for `pros` (choose one or more):**
> - Natural-Looking Result
> - Eliminates Desired Level Of Gray
> - Long-Lasting
> - Color Matches Desired Result
> - Conditioner Makes My Hair Feel Great
> - Brush Makes Application Easy
> - High Quality Gloves That Fit My Hands
> - Customized Instructions, Easy To Understand
> - Auto-Delivery
> - Mess-Free
> - Works Fast
>
> **Valid values for `cons` (choose one or more, or leave empty):**
> - Covered Too Much Gray
> - Covered Too Little Gray
> - Fades Quickly
> - Color Was Lighter Than I Expected
> - Color Was Darker Than I Expected
> - Difficult To Apply
> - Confusing Instructions
> - Messy
> - Time-Consuming
> - Stains In Skin / Sink
>
> **Important rules:**
> - Only select values that are clearly supported by the review text. Do not guess or assume.
> - If the comment is vague or doesn't mention any specific pros or cons, select the most general applicable option (e.g. "Natural-Looking Result" for a positive comment about the color).
> - Every review must have at least one `pro`.
> - Return the complete dataset as a downloadable CSV file, preserving all original columns and adding `pros` and `cons` as the last two columns.

---

Once you receive the enriched CSV from the AI tool, save it as **`verified_reviews_json_format.csv`** in the root of the repository.

---

### Step 3 – Upload Reviews – `upload_reviews_to_dev_legacy.py`

#### What the script does

- Reads **`verified_reviews_json_format.csv`** (output from Step 2)
- Connects to the `prod_ecommerce` MySQL database
- For each review row:
  - Generates a unique `REV...` review ID
  - Determines the `productReviewTypeId` from the SKU (beard kit, all-in-one, grooming tools, etc.)
  - Maps the SKU to a product image via `SKU_TO_PICTURE`
  - Normalizes the date, nickname, rating, and recommendation fields
  - Inserts the row into the `review` table with `visible = 0` (hidden until manually approved)
- Commits every 50 rows and prints progress

#### Required CSV columns

| Column | Description |
|---|---|
| `sku` | Product SKU |
| `overallrating` | Star rating (integer 1–5) |
| `date` | Review date (ISO 8601 format) |
| `headline` | Review title |
| `comment` | Full review text |
| `nickname` | Reviewer name (only first word is used) |
| `email` | Reviewer email |
| `recommendation` | Numeric recommendation score |
| `pros` | JSON array of pros (e.g. `["Natural-Looking Result"]`) |
| `cons` | JSON array of cons (e.g. `["Fades Quickly"]` or `[]`) |

#### How to run

```bash
python upload_reviews_to_dev_legacy.py
```

**Prerequisites:**
- `verified_reviews_json_format.csv` must exist in the root folder
- `.env` file must contain valid `DB_HOST`, `DB_USER`, and `DB_PASSWORD` values

> ⚠️ Reviews are inserted with **`visible = 0`**. After uploading, they must be manually set to visible (`visible = 1`) in the database or through the admin panel before they appear on the website.

### After uploading

1. **Upload the generated file** (`verified_reviews_json_format.csv` or the enriched Excel) to the following Drive folder to maintain a historical tracker of all uploaded reviews:
   [📁 Reviews Upload Tracker Folder](https://drive.google.com/drive/folders/1v2oyCe7vdGfuPviYzm96bQgbzrlDWLTZ?usp=drive_link)

2. **Share the Drive link** of the uploaded file in the **`#consumer-information-center`** Slack channel.

---

## 7. Repurchase & Cancellations Pipeline

This pipeline analyzes **first-order repurchase behavior** and **subscription cancellation reasons** by shade, ethnicity, and color experience. It consists of **three sequential steps** and is run **quarterly** (always for the previous quarter — see [Pipeline Frequency](#3-pipeline-frequency--schedule)).

---

### Step 1 – Repurchase Diagnostic – `repurchaseFirstOrderDiagnosticTotal.py`

#### What it does

- Queries the database to find customers who placed their **first order as an OTO (One-Time Order)** within the specified date range and whose order contained items from a specific product category
- For each customer, it retrieves their **diagnostic profile** (quiz answers at the time of purchase) and counts how many additional orders they placed afterward
- Analyzes repurchase rates broken down by every quiz variable (e.g. hair type, gray concentration, desired shade, experience with color, etc.)
- Generates an Excel file: **`analisis_recompra_consolidado.xlsx`** with the following sheets:
  - **`Todos los diagnósticos`** – repurchase rates for all diagnostic variables
  - **`Con Developer 20Vol`** – filtered to orders containing 20 Vol developer
  - **`Con Developer 10Vol`** – filtered to orders containing 10 Vol developer
  - **`Combinaciones`** – if combination mode is selected, analyzes repurchase by predefined variable combinations (e.g. experience with color + skin reaction)

#### Category filter (hair vs. beard)

Inside the `procesar_rango_fechas` function, the following SQL block controls which product category is used to filter first orders:

```sql
WHERE soi.salesOrderId = fo.id
    AND soi.category = 'IG00000000000000000000000000000029'
```

- `IG00000000000000000000000000000029` → **Hair** category
- `IG00000000000000000000000000000028` → **Beard** category

Change the category ID in this `WHERE` clause to switch between hair and beard analysis before running the script.

#### Date range configuration

Open the script and edit the `rangos_fechas` list in the `main()` function to set the desired quarter:

```python
rangos_fechas = [
    {'start': '2025-01-01', 'end': '2025-04-01', 'codigo': 'Q1 - 2025', 'nombre': 'Q1 - 2025'}
]
```

#### How to run

```bash
python repurchaseFirstOrderDiagnosticTotal.py
```

A GUI dialog will appear asking whether you want to run a **combination analysis** (`Yes`) or a **traditional diagnostic analysis** (`No`).

---

### Step 2 – Shade Cancellations – `shadeCancelations.py`

#### What it does

- Queries the database for all **subscription cancellations** within the specified date range, filtered to subscriptions that had a colorant shade item in their last order
- For each cancellation, retrieves the **cancellation reason**, the **last shade ordered**, and the customer's **diagnostic profile** (ethnicity and experience with color)
- Also queries **active subscriptions** in the same period for distribution context
- Generates an Excel file: **`analisis_cancelaciones_<startDate>_to_<endDate>.xlsx`** with three sheets:
  - **`Por Razon (Etnias)`** – cancellation reasons broken down by ethnicity (Caucasian / African / Asian)
  - **`Por Razon (Shades)`** – cancellation reasons broken down by the last shade ordered, with sub-tables per ethnicity
  - **`Por Razon (Experience)`** – cancellation reasons broken down by experience with color (Currently Dyed / I've colored / Never colored)

#### Date range configuration

Edit the `__main__` block at the bottom of the script:

```python
if __name__ == "__main__":
    main('2025-01-01', '2025-04-01')
```

#### How to run

```bash
python shadeCancelations.py
```

---

### Step 3 – Merge Report – `analisis_repurchase_cancelaciones.py`

#### What it does

This script takes the output files from Steps 1 and 2 and merges them into the master tracker file **`analisis_repurchase_cancelaciones.xlsx`**. It fills three sheets in the master file:

| Sheet | Source | Data |
|---|---|---|
| `Recompra` | `analisis_recompra_consolidado.xlsx` | Repurchase percentage per diagnostic variable |
| `Cancelaciones_Etnias` | `analisis_cancelaciones_*.xlsx` | Cancellation % by reason and ethnicity |
| `Cancelaciones_shade` | `analisis_cancelaciones_*.xlsx` | Cancellation % by reason and shade |

#### Configuration before running

Open the script and update the file names and column indexes at the top:

```python
RECOMPRA_FILE   = "analisis_recompra_consolidado.xlsx"
CANCEL_FILE     = "analisis_cancelaciones_2025-10-01_to_2026-01-01.xlsx"  # ← update this
MAESTRO_FILE    = "analisis_repurchase_cancelaciones.xlsx"

COL_IDX_RECOMPRA        = 8   # ← column index to write repurchase data in master
COL_IDX_CANCEL_ETNIAS   = 5   # ← column index to write ethnicity cancellation data
COL_IDX_CANCEL_SHADES   = 4   # ← column index to write shade cancellation data
```

Update `CANCEL_FILE` to match the filename generated by Step 2, and adjust the `COL_IDX_*` values to point to the correct quarter column in the master tracker.

#### How to run

```bash
python analisis_repurchase_cancelaciones.py
```

The script will print a summary of how many values were written to each sheet.

### Where to upload

After running the full pipeline, upload the results to the following shared Drive folder:

[📁 Repurchase & Cancellations Reports Folder](https://drive.google.com/drive/folders/1NM0jJ_zMBSRpaDdjt8nbHGIfHkneLCv0?usp=drive_link)

---

## 8. Supporting Modules

All shared modules live in the **`modules/`** directory.

### `modules/database_queries.py`

Provides the `execute_query(sql)` function that runs any SQL string against the configured MySQL database and returns a `pandas.DataFrame`.

### `modules/date_selector.py`

A Tkinter-based GUI that lets the user:
- Select start and end dates using an interactive calendar (`tkcalendar`)
- Enter the output folder name
- Toggle which report sections to include (orders, sales, payment errors, renewals, FC, subscriptions, refill, upsize, etc.)

Returns all selections to `main.py`.

### `modules/excel_creator.py`

Handles all Excel file generation using `openpyxl` and `matplotlib`:
- **`save_dataframe_to_excel`** – creates a `.xlsx` with tabular data and embedded line/bar charts
- **`line_chart`** – daily metric line chart saved as image and inserted into Excel
- **`bar_chart`** – comparative bar chart for new vs. existing users, recurring vs. non-recurring, etc.
- **`save_error_reasons_with_chart`** – writes payment error reasons with dynamic coloring and a bar chart
- **`save_dataframe_to_excel_orders`** – specialized sheet for order data with charts

### `modules/colors.py`

Provides `lighten_color(hex_color, factor=0.5)` – returns a lighter version of any hex color, used for chart styling.

---

## 9. Cloud Upload

### `uploadCloud.py`

Supports uploading generated report files to **Google Drive** or **Dropbox**.

#### Google Drive Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project and enable the **Google Drive API**
3. Download the **OAuth credentials JSON** file
4. Rename it to `credentials.json` and place it in the root of the repository
5. On first run, a browser window will open for you to authorize access

#### Dropbox Setup

1. Go to [Dropbox Developers](https://www.dropbox.com/developers/apps)
2. Create a new app and generate an **access token**
3. Open `uploadCloud.py` and paste the token on line 32:
   ```python
   dbx = dropbox.Dropbox('YOUR_TOKEN_HERE')
   ```

#### How It Works

When running `main.py`, a popup will appear after the reports are generated allowing you to select:
- ✅ **Dropbox** – uploads to `/MyReports/<folder_name>/<filename>`
- ✅ **Google Drive** – uploads to the configured folder ID
- (Neither selected) – files are kept locally only

---

## 10. General Metrics Report

The file **`metricas.xlsx`** (provided in the repository) serves as a centralized tracker where all monthly metrics are recorded.

- `report.py` contains `anotar_datos_excel()`, which writes extracted values into the correct rows and columns of `metricas.xlsx`
- The column (month) is determined by the `columna` variable in `main.py` (default: 18)
- If you rename the Excel file or the sheet, update the `archivo_excel` and `hoja_nombre` variables in `report.py`
- This step is optional – if you do not need the centralized tracker, you can delete `metricas.xlsx` and the script will still generate all individual report files

---

## 11. Project Structure

```
Cleverman-Metrics/
├── .env                          # Database credentials (not committed)
├── credentials.json              # Google Drive OAuth credentials (not committed)
├── metricas.xlsx                 # Centralized monthly metrics tracker
├── reviews.json                  # Raw reviews JSON from SUVAE (input for read_reviews.py)
│
├── main.py                       # ⭐ Main entry point – Monthly Report
├── fcReport.py                   # ⭐ Full Control Report
├── read_reviews.py               # ⭐ Step 1 – Parse reviews JSON
├── upload_reviews_to_dev_legacy.py # ⭐ Step 3 – Upload reviews to DB
│
├── repurchaseFirstOrderDiagnosticTotal.py  # ⭐ Repurchase pipeline – Step 1
├── shadeCancelations.py                    # ⭐ Repurchase pipeline – Step 2
├── analisis_repurchase_cancelaciones.py    # ⭐ Repurchase pipeline – Step 3
│
├── modules/
│   ├── database_queries.py       # Shared DB query helper
│   ├── date_selector.py          # GUI date/option selector
│   ├── excel_creator.py          # Excel & chart generation
│   └── colors.py                 # Color utilities
│
├── orders.py                     # Order metrics
├── payments.py                   # Payment error metrics
├── renewalsAndNoRecurrents.py    # Sales & renewals metrics
├── subscriptions.py              # Subscription breakdown
├── refill.py                     # Refill orders
├── upsize.py                     # Upsell metrics
├── exceptedRenewals.py           # Renewal forecasting
├── realRenewalFrecuency.py       # Renewal frequency
├── fullContol.py                 # FC wrapper (calls fcReport logic)
├── howHearFromUs.py              # Attribution survey
├── ga4Funnels.py                 # GA4 funnel CSV processor
├── selectFiles.py                # GUI file selector for CSVs/Stripe files
├── report.py                     # Writes metrics to metricas.xlsx
├── uploadCloud.py                # Google Drive & Dropbox upload
├── block_payments.py             # Stripe blocked payments analysis
├── aov_free_shipping.py          # AOV & free shipping analysis
├── backupPayment.py              # Backup payment analysis
├── backupPaymentMethod.py        # Backup payment method analysis
├── colorCancellations.py         # Color-based cancellation analysis
├── midBrownCancellations.py      # Mid-brown cancellation analysis
├── shadeBeardOrHairCancelations.py # Beard/hair shade cancellations
├── newRealRenewalFrecuency.py    # Updated renewal frequency analysis
└── 3x2promo.py                   # 3x2 promotion analysis
```

---

## Quick Reference – Which script to run for each task

| Task | Script | Frequency |
|---|---|---|
| Generate the full monthly report | `python main.py` | Monthly (beginning of month) |
| Generate the Full Control tracker | `python fcReport.py` | Monthly (beginning of month) |
| Parse reviews JSON from SUVAE | `python read_reviews.py` | Every 2 weeks |
| Upload enriched reviews to DB | `python upload_reviews_to_dev_legacy.py` | Every 2 weeks |
| Repurchase diagnostic analysis | `python repurchaseFirstOrderDiagnosticTotal.py` | Quarterly (previous quarter) |
| Shade cancellations analysis | `python shadeCancelations.py` | Quarterly (previous quarter) |
| Merge repurchase + cancellations report | `python analisis_repurchase_cancelaciones.py` | Quarterly (after steps 1 & 2) |
