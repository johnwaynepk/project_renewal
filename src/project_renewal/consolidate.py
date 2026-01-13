from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


TIERS_ALLOWED = {"enterprise", "enterprise ai", "inspect pro", "starter+"}
PRODUCT_ALLOWED = "inspect"


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
      - Tier in {Enterprise, Inspect Pro, Starter+} (case-insensitive)
    """
    required_cols = {"Contract ID", "Product", "Tier"}
    missing = required_cols - set(license_df.columns)
    if missing:
        raise ValueError(f"license.csv is missing columns: {sorted(missing)}")

    product_norm = license_df["Product"].astype(str).str.strip().str.lower()
    tier_norm = license_df["Tier"].astype(str).str.strip().str.lower()

    mask = product_norm.eq(PRODUCT_ALLOWED) & tier_norm.isin(TIERS_ALLOWED)
    return license_df.loc[mask].copy()


def consolidate(data_dir: Path) -> Path:
    contract_path = data_dir / "contract.csv"
    license_path = data_dir / "license.csv"
    finance_path = data_dir / "finance.csv"
    output_path = data_dir / "consolidated.csv"

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

    # Ensure the data folder exists (in case someone runs this on a fresh repo)
    data_dir.mkdir(parents=True, exist_ok=True)

    out_df.to_csv(output_path, index=False, encoding="utf-8")

    # Optional console summary (helpful for quick sanity check)
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
