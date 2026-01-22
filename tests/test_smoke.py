from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

import pandas as pd
import pytest

from project_renewal.consolidate import (
    consolidate,
    filter_license_rows,
    read_license_csv,
    read_standard_csv,
)

TESTS_DIR = Path(__file__).parent


@pytest.fixture
def test_data_dir(tmp_path: Path) -> Path:
    """Copy test data files to a temp directory for isolation."""
    for filename in ("contracts.csv", "licenses.csv", "finance.csv"):
        src = TESTS_DIR / filename
        if src.exists():
            shutil.copy(src, tmp_path / filename)
    return tmp_path


def test_consolidate_runs_without_error(test_data_dir: Path) -> None:
    """Smoke test: consolidate() should run without raising exceptions."""
    output_path = consolidate(test_data_dir)
    assert output_path.exists()
    assert output_path.suffix == ".xlsx"


def test_consolidate_output_has_expected_rows(test_data_dir: Path) -> None:
    """Verify filtering: only valid licenses should appear in output."""
    consolidate(test_data_dir)

    # Read the output Excel file
    output_path = test_data_dir / "consolidated.xlsx"
    df = pd.read_excel(output_path)

    # Test data has 8 license rows:
    # - 4 should pass (contract-001 to 004: valid Product, Tier, BU)
    # - 4 should be filtered out:
    #   - contract-005: BU = "Proceq HQ" (excluded)
    #   - contract-006: Product = "GPR Live" (not Inspect)
    #   - contract-007: BU = "" (empty, excluded)
    #   - contract-008: Tier = "Inspect Free Trial" (not allowed)
    assert len(df) == 4


def test_filter_license_rows() -> None:
    """Test the filter_license_rows function directly."""
    license_path = TESTS_DIR / "licenses.csv"
    df = read_license_csv(license_path)

    filtered = filter_license_rows(df)

    # Should keep only 4 rows with valid Product, Tier, and BU
    assert len(filtered) == 4

    # Verify the correct contracts were kept
    kept_ids = set(filtered["Contract ID"].tolist())
    assert kept_ids == {"contract-001", "contract-002", "contract-003", "contract-004"}


def test_contract_data_merged(test_data_dir: Path) -> None:
    """Verify contract fields are merged into output."""
    consolidate(test_data_dir)

    output_path = test_data_dir / "consolidated.xlsx"
    df = pd.read_excel(output_path)

    # Check that contract fields were added
    expected_cols = {"Country Sold To", "User Type", "Remarks", "BP", "Language"}
    assert expected_cols.issubset(set(df.columns))

    # Verify a specific contract's data was merged
    row = df[df["Contract ID"] == "contract-001"].iloc[0]
    assert row["Country Sold To"] == "Singapore"
    assert row["User Type"] == "End User"


def test_output_sorted_by_expiration(test_data_dir: Path) -> None:
    """Verify output is sorted by expiration date (earliest first)."""
    consolidate(test_data_dir)

    output_path = test_data_dir / "consolidated.xlsx"
    df = pd.read_excel(output_path)

    expirations = pd.to_datetime(df["Expiration"]).tolist()
    assert expirations == sorted(expirations)
