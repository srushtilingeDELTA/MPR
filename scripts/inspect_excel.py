"""Inspect Excel structure: sheets, header rows, column names."""

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"

# Try common filenames
candidates = list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.xlsm"))
if not candidates:
    print(f"No Excel files in {DATA_DIR}")
    raise SystemExit(1)

file_path = candidates[0]
print(f"File: {file_path.name}\n")

xl = pd.ExcelFile(file_path, engine="openpyxl")
print("Sheet names:")
for i, name in enumerate(xl.sheet_names):
    print(f"  [{i}] {name}")

print("\n" + "=" * 60)
for sheet in xl.sheet_names:
    print(f"\n--- Sheet: {sheet} ---")
    preview = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=8, engine="openpyxl")
    print("First 8 rows (raw, no header):")
    print(preview.to_string())

    for header_row in range(0, 5):
        df = pd.read_excel(file_path, sheet_name=sheet, header=header_row, nrows=3, engine="openpyxl")
        cols = [str(c).strip() for c in df.columns]
        print(f"\n  If header is row {header_row + 1}, columns are:")
        print(f"  {cols}")
