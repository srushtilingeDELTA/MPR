from pathlib import Path
import pandas as pd
import yaml

BASE = Path(__file__).resolve().parent.parent
config = yaml.safe_load(open(BASE / "config.yaml", encoding="utf-8"))
path = BASE / config["excel"]["file_path"]
if not path.exists():
    print(f"Missing: {path}")
else:
    xl = pd.ExcelFile(path, engine="openpyxl")
    print("Sheet names:")
    for n in xl.sheet_names:
        print(f"  - {n}")
        