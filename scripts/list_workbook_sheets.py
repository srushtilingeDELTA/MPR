"""List sheets in all configured workbooks."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from mpr_data import MprData
from report_utils import load_config


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    data = MprData(config, BASE_DIR)
    data.load()
    for name in config.get("workbooks", {}):
        print(f"\nWorkbook: {name}")
        for sheet in data.sheet_names(name):
            print(f"  - {sheet}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
