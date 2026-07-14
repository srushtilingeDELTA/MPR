"""List workbook sheets and suggested screenshot mappings for template slides."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config
from workbook_store import WorkbookStore


SUGGESTED = [
    (2, "System sections: Safety/Customer/Ops"),
    (3, "System sections: People/Finance/Overall"),
    (4, "GIR native tables + chart from Workings!GIR"),
    (5, "workings EAC/ASAP single table screenshot"),
    (6, "workings PEOPLE table + 3 Excel graph screenshots"),
    (7, "workings FINANCE Regions/BUDGET/OVERTIME/TOTAL HOURS screenshot"),
    (10, "workings PMI Regions Mot/Stat+Non-Mot table + Mot/Stat graphs"),
    (11, "workings ISR Regions Rel/Sev table + Rel/Sev graphs"),
    (13, "North summary"),
    (14, "North comparison"),
    (15, "South summary"),
    (16, "South comparison"),
    (17, "Stationary summary"),
    (18, "Stationary comparison"),
    (23, "Galley summary"),
    (24, "Galley comparison"),
]


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    store = WorkbookStore(config, BASE_DIR)
    try:
        store.load()
    except FileNotFoundError as exc:
        print(exc)
        return 1

    for key in ("scorecards", "workings", "actuals"):
        names = store.sheet_names(key)
        print(f"\n{key} ({len(names)} sheets):")
        if not names:
            print("  (not loaded)")
            continue
        for idx, name in enumerate(names):
            print(f"  [{idx}] {name}")

    print("\nSuggested screenshot slide mapping:")
    for slide_idx, note in SUGGESTED:
        print(f"  PPT slide {slide_idx + 1} (index {slide_idx}): {note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
