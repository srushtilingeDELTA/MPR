"""List sheet names in the MPR Excel workbook."""

from pathlib import Path

import pandas as pd
import yaml

BASE_DIR = Path(__file__).resolve().parent.parent


def main() -> None:
    with open(BASE_DIR / "config.yaml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    file_path = BASE_DIR / config["excel"]["file_path"]
    if not file_path.exists():
        print(f"File not found: {file_path}")
        print("Download the Excel file from SharePoint into the data/ folder.")
        return

    xl = pd.ExcelFile(file_path, engine="openpyxl")
    print("Sheet names:")
    for name in xl.sheet_names:
        print(f"  - {name}")
    print(f"\nCurrent config sheet_name: {config['excel']['sheet_name']}")


if __name__ == "__main__":
    main()
