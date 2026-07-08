"""Test live SharePoint Excel read (Edge login + in-memory download)."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config


def main() -> None:
    config = load_config(base_dir=BASE_DIR)
    from sharepoint_live import fetch_workbook_bytes

    data = fetch_workbook_bytes(config, base_dir=BASE_DIR)
    print(f"OK: fetched {len(data):,} bytes from SharePoint (in memory)")

    import pandas as pd

    xl = pd.ExcelFile(__import__("io").BytesIO(data), engine="openpyxl")
    print(f"Sheets: {len(xl.sheet_names)}")
    sheet = config["excel"]["sheet_name"]
    df = pd.read_excel(xl, sheet_name=sheet, header=config["excel"].get("header_row", 0))
    print(f"{sheet}: {len(df)} rows x {len(df.columns)} columns")


if __name__ == "__main__":
    main()