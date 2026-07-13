"""List unique KPI names in the report month sheet."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import pandas as pd

from report_utils import load_config


def main() -> None:
    config = load_config(base_dir=BASE_DIR)
    cols = config["columns"]
    path = BASE_DIR / config["excel"]["file_path"]
    if not path.exists():
        print(f"Excel not found: {path}")
        return

    df = pd.read_excel(
        path,
        sheet_name=config["excel"]["sheet_name"],
        header=config["excel"].get("header_row", 0),
        engine="openpyxl",
    )
    df.columns = df.columns.astype(str).str.strip()
    kpi_col = cols["kpi"]
    print(f"Sheet: {config['excel']['sheet_name']}")
    print(f"Rows: {len(df)}")
    print("\nUnique KPI values:")
    for kpi in sorted(df[kpi_col].dropna().astype(str).unique()):
        count = len(df[df[kpi_col].astype(str) == kpi])
        print(f"  - {kpi} ({count} rows)")


if __name__ == "__main__":
    main()
