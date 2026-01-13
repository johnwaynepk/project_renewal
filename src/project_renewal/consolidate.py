from __future__ import annotations
from openpyxl.styles import PatternFill

import argparse
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


TIERS_ALLOWED = {"enterprise", "enterprise ai", "inspect pro", "starter+"}
PRODUCT_ALLOWED = "inspect"
BU_EXCLUDE = "proceq hq"
BU_MISSING_MARKERS = {"", "#n/a", "n/a", "na", "none", "null"}

ORANGE_FILL = PatternFill(start_color="FFFFA500", end_color="FFFFA500", fill_type="solid")  # Orange
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")  # Yellow

def project_root_from_this_file() -> Path:
    """
    Returns the project root, assuming this file lives at:
      <root>/src/project_renewal/consolidate.py
    """
    return Path(__file__).resolve().parents[2]


def read_license_csv(path: Path) -> pd.DataFrame:
    """
    license.csv in your sample has an extra first line like:
        Contract,,,,,License,,,,,,
    so we skip that row (skiprows=1) and read the real header on line 2.

    We also use keep_default_na=False so strings like "NA" are not turned into NaN.
    """
    df = pd.read_csv(
        path,
        encoding="utf-8-sig",
        skiprows=1,
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [c.strip() for c in df.columns]
    return df


def read_standard_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(
        path,
        encoding="utf-8-sig",
        dtype=str,
        keep_default_na=False,
    )
    df.columns = [c.strip() for c in df.columns]
    return df


def build_contract_lookup(contract_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    """
    Build a dictionary:
      contract_lookup[contract_id] = {
        "Country Sold To": ...,
        "User Type": ...,
        "Remarks": ...,
        "BP": ...,
        "First Expiration Date": ...,
        "License Count": ...,
        "Language": ...,
      }
    Where contract_id is contract_df["ID"].

    If duplicate IDs exist, the first occurrence is used.
    """
    required_cols = {"ID", "Country Sold To", "User Type", "Remarks", "BP", "First Expiration Date", "License Count", "Language"}
    missing = required_cols - set(contract_df.columns)
    if missing:
        raise ValueError(f"contract.csv is missing columns: {sorted(missing)}")

    fields = [
        "Country Sold To",
        "User Type",
        "Remarks",
        "BP",
        "First Expiration Date",
        "License Count",
        "Language",
    ]

    lookup: Dict[str, Dict[str, str]] = {}
    for _, row in contract_df.iterrows():
        cid = str(row["ID"]).strip()
        if not cid or cid in lookup:
            continue
        lookup[cid] = {f: str(row.get(f, "")).strip() for f in fields}
    return lookup


def build_finance_lookup(finance_df: pd.DataFrame) -> Dict[str, str]:
    """
    Build a dictionary:
      finance_lookup[contract_id] = customer_name

    If multiple rows exist for the same Contract ID, we take the first non-empty Customer name.
    If none is found, it will not be present in the dict (caller will default to "#NA").
    """
    required_cols = {"Contract ID", "Customer name"}
    missing = required_cols - set(finance_df.columns)
    if missing:
        raise ValueError(f"finance.csv is missing columns: {sorted(missing)}")

    lookup: Dict[str, str] = {}
    for _, row in finance_df.iterrows():
        cid = str(row["Contract ID"]).strip()
        cname = str(row.get("Customer name", "")).strip()
        if not cid:
            continue

        # If we already have a non-empty name for this CID, keep it.
        if cid in lookup and lookup[cid] != "#NA":
            continue

        if cname:
            lookup[cid] = cname
        else:
            # Only set #NA if nothing set yet
            lookup.setdefault(cid, "#NA")

    return lookup


def filter_license_rows(license_df: pd.DataFrame) -> pd.DataFrame:
    """
    Filter license.csv:
      - Product == Inspect (case-insensitive)
      - Tier in {Enterprise, Enterprise AI, Inspect Pro, Starter+} (case-insensitive)
      - BU is NOT:
          * "Proceq HQ"
          * blank/empty
          * missing markers like "#N/A", "N/A", "NA", "None", "Null"
    """
    required_cols = {"Contract ID", "Product", "Tier", "BU"}
    missing = required_cols - set(license_df.columns)
    if missing:
        raise ValueError(f"license.csv is missing columns: {sorted(missing)}")

    product_norm = license_df["Product"].astype(str).str.strip().str.lower()
    tier_norm = license_df["Tier"].astype(str).str.strip().str.lower()
    bu_norm = license_df["BU"].astype(str).str.strip().str.lower()

    mask = (
        product_norm.eq(PRODUCT_ALLOWED)
        & tier_norm.isin(TIERS_ALLOWED)
        & ~bu_norm.eq(BU_EXCLUDE)
        & ~bu_norm.isin(BU_MISSING_MARKERS)
    )
    return license_df.loc[mask].copy()

def autofit_excel_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Auto-fit column widths in the given Excel sheet based on cell content length.
    Works with openpyxl engine.
    """
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns, 1):
        # Compute max width from column header + all values (converted to string)
        series = df[col].astype(str).fillna("")
        max_len = max([len(str(col))] + series.map(len).tolist())
        worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 60)

def highlight_expiring_rows(
    writer: pd.ExcelWriter,
    sheet_name: str,
    df: pd.DataFrame,
    date_col: str = "Expiration",
) -> None:
    """
    Highlight entire rows based on Expiration date:
      - 0–30 days from today: orange
      - 31–60 days from today: yellow

    Rows with missing/invalid dates (blank, #N/A, NA, etc.) are not highlighted.
    """
    if date_col not in df.columns:
        return

    ws = writer.sheets[sheet_name]

    # Parse dates robustly (strings -> datetime). Invalid values become NaT.
    exp = pd.to_datetime(df[date_col], errors="coerce")

    today = pd.Timestamp.today().normalize()

    # Excel row indexing:
    # row 1 = header, row 2 = first data row
    for excel_row_idx, exp_ts in enumerate(exp, start=2):
        if pd.isna(exp_ts):
            continue

        days_to_expiry = (exp_ts.normalize() - today).days

        if 0 <= days_to_expiry <= 30:
            fill = ORANGE_FILL
        elif 31 <= days_to_expiry <= 60:
            fill = YELLOW_FILL
        else:
            continue

        # Apply fill across the whole row (all columns in df)
        for col_idx in range(1, df.shape[1] + 1):
            ws.cell(row=excel_row_idx, column=col_idx).fill = fill


def consolidate(data_dir: Path) -> Path:
    contract_path = data_dir / "contract.csv"
    license_path = data_dir / "license.csv"
    finance_path = data_dir / "finance.csv"
    output_path = data_dir / "consolidated.xlsx"

    # ---- Read inputs ----
    lic_df = read_license_csv(license_path)
    con_df = read_standard_csv(contract_path)
    fin_df = read_standard_csv(finance_path)

    # ---- Filter licenses ----
    lic_filtered = filter_license_rows(lic_df)

    # ---- Create data structure (list of dicts) ----
    items: List[Dict[str, Any]] = lic_filtered.to_dict(orient="records")

    # ---- Build lookups ----
    contract_lookup = build_contract_lookup(con_df)
    finance_lookup = build_finance_lookup(fin_df)

    # ---- Enrich each item (by Contract ID) ----
    contract_fields = [
        "Country Sold To",
        "User Type",
        "Remarks",
        "BP",
        "First Expiration Date",
        "License Count",
        "Language",
    ]

    missing_contract = 0
    missing_finance = 0

    for item in items:
        contract_id = str(item.get("Contract ID", "")).strip()

        # Contract fields: match license["Contract ID"] -> contract["ID"]
        cinfo = contract_lookup.get(contract_id)
        if cinfo is None:
            missing_contract += 1
            # Not specified to use #NA here, so leave blanks for contract fields.
            for f in contract_fields:
                item.setdefault(f, "")
        else:
            item.update(cinfo)

        # Finance: match license["Contract ID"] -> finance["Contract ID"]
        customer_name = finance_lookup.get(contract_id, "#NA")
        if customer_name == "#NA":
            missing_finance += 1
        item["Customer name"] = customer_name

    # ---- Output ----
    out_df = pd.DataFrame(items)

    # Use "Expiration" as the key: sort by Expiration date (invalid/missing go last)
    if "Expiration" in out_df.columns:
        _exp_sort = pd.to_datetime(out_df["Expiration"], errors="coerce")
        out_df = (
            out_df.assign(_exp_sort=_exp_sort)
            .sort_values("_exp_sort", na_position="last")
            .drop(columns=["_exp_sort"])
        )

    # ---- Ensure "Customer name" is the 2nd column ----
    if "Customer name" not in out_df.columns:
        out_df["Customer name"] = "#NA"

    cols = list(out_df.columns)

    # Prefer Contract ID as the first column if present; otherwise keep the first non-customer-name column
    if "Contract ID" in cols:
        first_col = "Contract ID"
    else:
        first_col = next((c for c in cols if c != "Customer name"), cols[0])

    new_cols = [first_col, "Customer name"] + [c for c in cols if c not in {first_col, "Customer name"}]
    out_df = out_df[new_cols]

    data_dir.mkdir(parents=True, exist_ok=True)

    sheet_name = "consolidated"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Highlight expiring rows based on Expiration:
        highlight_expiring_rows(writer, sheet_name, out_df, date_col="Expiration")

        # If you already have autofit_excel_columns(), keep using it:
        autofit_excel_columns(writer, sheet_name, out_df)

    print(f"Filtered licenses: {len(lic_filtered):,} rows")
    print(f"Missing contract matches (ID not found in contract.csv): {missing_contract:,}")
    print(f"Missing finance matches (Customer name defaulted to #NA): {missing_finance:,}")
    print(f"Output written to: {output_path}")

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Consolidate contract/license/finance CSVs into data/consolidated.csv"
    )
    parser.add_argument(
        "--data-dir",
        type=str,
        default=None,
        help="Path to the data folder (default: <project_root>/data).",
    )
    args = parser.parse_args()

    root = project_root_from_this_file()
    data_dir = Path(args.data_dir) if args.data_dir else (root / "data")

    consolidate(data_dir)


if __name__ == "__main__":
    main()
