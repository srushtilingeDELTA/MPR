"""Inspect System tab sections in 2026 - GSE Scorecards.xlsx.

Usage:
    python scripts/inspect_scorecard_system.py
"""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from openpyxl.utils import get_column_letter

from report_utils import load_config
from scorecard_screenshots import detect_system_layout, select_sections_for_slide
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

    layout = detect_system_layout(raw, sheet_name="System")
    sections = layout.sections
    if not sections:
        print("No sections detected on System sheet.")
        return 1

    print(
        f"Header rows {layout.header_start_row}-{layout.header_end_row}  "
        f"cols {get_column_letter(layout.start_col)}-{get_column_letter(layout.end_col)}"
    )
    print(f"Found {len(sections)} category section(s) on System:\n")
    for section in sections:
        flag = " [BLACK]" if section.is_black else ""
        print(
            f"  {section.index + 1}. {section.title!r}  "
            f"rows {section.start_row}-{section.end_row}{flag}"
        )

    slide3 = select_sections_for_slide(sections, mode="first", count=3)
    slide4 = select_sections_for_slide(sections, mode="last", count=2, include_black=True)
    b3 = layout.capture_bounds_for(slide3)
    b4 = layout.capture_bounds_for(slide4)
    print("\nSlide 3 (PPT) capture:")
    print(
        f"  {_fmt(layout.header_start_row, b3[1], b3[2], b3[3])}  -> "
        + ", ".join(s.title for s in slide3)
    )
    print("\nSlide 4 (PPT) capture:")
    if layout.needs_header_stitch(slide4):
        print(
            f"  header {_fmt(layout.header_start_row, layout.header_end_row, b4[2], b4[3])} "
            f"+ body {_fmt(b4[0], b4[1], b4[2], b4[3])}  -> "
            + ", ".join(s.title for s in slide4)
        )
    else:
        print(
            f"  {_fmt(layout.header_start_row, b4[1], b4[2], b4[3])}  -> "
            + ", ".join(s.title for s in slide4)
        )
    return 0


def _fmt(r1: int, r2: int, c1: int, c2: int) -> str:
    return f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"


if __name__ == "__main__":
    raise SystemExit(main())
