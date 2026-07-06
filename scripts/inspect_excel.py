from pathlib import Path
import pandas as pd

BASE = Path(__file__).resolve().parent.parent
files = list((BASE / "data").glob("*.xlsx")) + list((BASE / "data").glob("*.xlsm"))
path = files[0]
print(f"File: {path.name}\n")

xl = pd.ExcelFile(path, engine="openpyxl")
print("Sheets:")
for i, s in enumerate(xl.sheet_names):
    print(f"  [{i}] {s}")

for sheet in xl.sheet_names:
    print(f"\n--- {sheet} ---")
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=6, engine="openpyxl")
    print(raw.to_string())
    for hr in range(5):
        df = pd.read_excel(path, sheet_name=sheet, header=hr, nrows=2, engine="openpyxl")
        cols = [str(c).strip() for c in df.columns]
        print(f"  header row {hr+1}: {cols}")