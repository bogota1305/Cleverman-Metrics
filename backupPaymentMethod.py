"""
backup_payment_methods_report.py

Lee un CSV con intentos de pago, agrupa por (Description + Customer Email),
detecta errores, resoluciones, y resoluciones por "backup payment method",
y exporta un Excel con:
- Sheet 1: "Backup Resolved" (solo pagos resueltos por backup)
- Sheet 2: "Summary" (totales + porcentajes)

Uso:
  python backup_payment_methods_report.py --input payments.csv --output report.xlsx

Requisitos:
  pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
from pathlib import Path
import pandas as pd
import sys


REQUIRED_COLUMNS = [
    "Description",
    "Customer Email",
    "Status",
    "Created date (UTC)",
    "Card ID",
]


def normalize_status(s: str) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def safe_pct(n: int, d: int) -> float:
    return (n / d) if d else 0.0


def build_report(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Returns:
      detail_df: 1 row per payment group
      backup_resolved_df: subset of detail_df where Resolved by Backup == Yes
      summary_df: metrics
    """
    # Parse date
    df["Created date (UTC)"] = pd.to_datetime(
        df["Created date (UTC)"],
        format="%Y-%m-%d %H:%M:%S",
        errors="coerce",
    )

    # Basic validation
    if df["Created date (UTC)"].isna().any():
        bad = df[df["Created date (UTC)"].isna()][
            ["Created date (UTC)", "Customer Email", "Description"]
        ].head(10)
        raise ValueError(
            "Some rows have invalid 'Created date (UTC)' (expected format YYYY-MM-DD HH:MM:SS). "
            f"Examples (up to 10):\n{bad.to_string(index=False)}"
        )

    # Normalize Status
    df["_status_norm"] = df["Status"].map(normalize_status)

    group_cols = ["Customer Email", "Description"]

    records: list[dict] = []
    df_sorted = df.sort_values("Created date (UTC)")

    for (email, desc), g in df_sorted.groupby(group_cols, sort=False):
        g = g.sort_values("Created date (UTC)")
        attempts = int(len(g))

        had_error = (g["_status_norm"] == "failed").any()

        # "Resolved": had errors and last attempt ended in Paid
        last_status = g.iloc[-1]["_status_norm"]
        resolved = bool(had_error and last_status == "paid")

        failed_ids = set(
            g.loc[g["_status_norm"] == "failed", "Card ID"].astype(str).tolist()
        )
        paid_ids = g.loc[g["_status_norm"] == "paid", "Card ID"].astype(str).tolist()
        paid_id = paid_ids[-1] if len(paid_ids) > 0 else None

        # Resolved by backup:
        # - resolved == True
        # - all failed attempts used the same Card ID (len(failed_ids) == 1)
        # - final paid attempt uses a different Card ID than the failed one
        resolved_by_backup = False
        if resolved and len(failed_ids) == 1 and paid_id is not None:
            only_failed_id = next(iter(failed_ids))
            resolved_by_backup = (paid_id != only_failed_id)

        records.append(
            {
                "Customer Email": email,
                "Description": desc,
                "Attempts": attempts,
                "Had Error": "Yes" if had_error else "No",
                "Resolved": "Yes" if resolved else "No",
                "Resolved by Backup": "Yes" if resolved_by_backup else "No",
                "Failed Card IDs": (
                    "{" + ", ".join(sorted(failed_ids)) + "}" if failed_ids else "{}"
                ),
                "Paid Card ID": paid_id if paid_id is not None else "(none)",
                "First Attempt UTC": g["Created date (UTC)"].min(),
                "Last Attempt UTC": g["Created date (UTC)"].max(),
            }
        )

    detail_df = pd.DataFrame(records)

    total_payments = int(len(detail_df))
    total_errors = int((detail_df["Had Error"] == "Yes").sum())
    total_resolved = int((detail_df["Resolved"] == "Yes").sum())
    total_resolved_backup = int((detail_df["Resolved by Backup"] == "Yes").sum())

    # IMPORTANT:
    # El usuario pidió estos campos y estas fórmulas (hay dos labels muy parecidos, pero respetamos su fórmula):
    # - porcentaje de errores de pago = total pagos resueltos / total errores de pago
    # - porcentaje de pagos resueltos = total pagos resueltos / total errores de pago
    # (son iguales según lo que escribiste; si quieres cambiar uno luego, lo ajustamos)
    summary_df = pd.DataFrame(
        [
            {"Metric": "Total payments", "Value": total_payments},
            {
                "Metric": "Total payment errors (payments with ≥1 Failed)",
                "Value": total_errors,
            },
            {"Metric": "Total resolved payments (error resolved)", "Value": total_resolved},
            {"Metric": "Total resolved by backup", "Value": total_resolved_backup},
            {
                "Metric": "Porcentaje de errores de pago (total errors/ total payments)",
                "Value": safe_pct(total_errors, total_payments),
            },
            {
                "Metric": "Porcentaje de pagos resueltos (total resolved / total errors)",
                "Value": safe_pct(total_resolved, total_errors),
            },
            {
                "Metric": "Porcentaje de resueltos por backup (resolved by backup / total resolved)",
                "Value": safe_pct(total_resolved_backup, total_resolved),
            },
        ]
    )

    backup_resolved_df = detail_df[detail_df["Resolved by Backup"] == "Yes"].copy()

    return detail_df, backup_resolved_df, summary_df


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate backup payment method usage report from a CSV export."
    )
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="Path to input CSV file.",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="backup_payment_method_report.xlsx",
        help="Path to output Excel file (.xlsx).",
    )
    parser.add_argument(
        "--encoding",
        default=None,
        help="Optional CSV encoding (e.g. utf-8, latin1). If omitted, pandas default is used.",
    )
    if len(sys.argv) == 1:
        sys.argv.extend([
            "-i", "payments.csv",
            "-o", "backup_payment_report.xlsx"
        ])

    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        raise FileNotFoundError(f"Input CSV not found: {in_path}")

    # Read CSV
    df = pd.read_csv(in_path, encoding=args.encoding)

    # Validate required columns
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Missing required columns in CSV: "
            + ", ".join(missing)
            + "\nRequired columns are: "
            + ", ".join(REQUIRED_COLUMNS)
        )

    # Build report
    detail_df, backup_resolved_df, summary_df = build_report(df)

    # Write Excel
    with pd.ExcelWriter(out_path, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss") as writer:
        # Sheet requested: only backup-resolved payments
        backup_resolved_df.to_excel(writer, sheet_name="Backup Resolved", index=False)

        # Summary sheet
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # (Opcional) Si quieres conservar el detalle completo, descomenta esto:
        # detail_df.to_excel(writer, sheet_name="All Payments (Detail)", index=False)

    print(f"✅ Report written to: {out_path.resolve()}")
    print("Sheets:")
    print(" - Backup Resolved")
    print(" - Summary")


if __name__ == "__main__":
    main()
