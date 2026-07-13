"""Inspect System tab sections in 2026 - GSE Scorecards.xlsx.

Usage:
    python scripts/inspect_scorecard_system.py
"""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config
from scorecard_screenshots import detect_system_sections, select_sections_for_slide
from workbook_store import WorkbookStore


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    store = WorkbookStore(config, BASE_DIR)
    try:
        store.load()
        raw = store.workbook_bytes("scorecards")
    except FileNotFoundError as exc:
        print(exc)
        print("Sync SharePoint files or place the workbook under data/ first.")
        return 1

    sections = detect_system_sections(raw, sheet_name="System")
    if not sections:
        print("No sections detected on System sheet.")
        return 1

    print(f"Found {len(sections)} section(s) on System:\n")
    for section in sections:
        flag = " [BLACK]" if section.is_black else ""
        print(f"  {section.index + 1}. {section.title!r}  {section.range_address}{flag}")

    slide3 = select_sections_for_slide(sections, mode="first", count=3)
    slide4 = select_sections_for_slide(sections, mode="last", count=2, include_black=True)
    print("\nSlide 3 (PPT) would use:")
    for section in slide3:
        print(f"  - {section.title!r} ({section.range_address})")
    print("\nSlide 4 (PPT) would use:")
    for section in slide4:
        print(f"  - {section.title!r} ({section.range_address})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
