"""Finance slide (PPT 8) from New GSE MPR Workings.xlsx → FINANCE tab.

Screenshots the Regions scorecard block:
  Regions | BUDGET $000s | OVERTIME | TOTAL HOURS
  (MTD / YTD / Score headers + Actual / (B)/W Goal data rows)
"""

from __future__ import annotations

import io
import logging
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from scorecard_screenshots import (
    _find_sheet_name,
    _range_address,
    _validate_capture,
    capture_range_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Template picture slot on slide 8 (from GSE MPR template).
FINANCE_TABLE_BOX = (206_375, 929_805, 11_779_250, 4_457_700)

_END_TOKENS = (
    "notes",
    "leading issues",
    "action plan",
    "action plans",
    "comments",
)


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row, col).value
    if val is None:
        return ""
    return str(val).strip()


def _row_blob(ws, row: int, max_col: int) -> str:
    return " ".join(_cell_str(ws, row, c).casefold() for c in range(1, max_col + 1))


def _find_finance_header(ws, *, max_row: int = 40, max_col: int = 40) -> tuple[int, int, int] | None:
    """Locate the Regions / BUDGET / OVERTIME / TOTAL HOURS header row.

    Returns (start_row, start_col, end_col) for the group-title row.
    """
    scan_cols = min(int(ws.max_column or max_col), max_col)
    for row in range(1, min(int(ws.max_row or max_row), max_row) + 1):
        blob = _row_blob(ws, row, scan_cols)
        if "regions" not in blob:
            continue
        if "budget" not in blob:
            continue
        if "overtime" not in blob and "total hours" not in blob:
            continue

        start_col = None
        end_col = 1
        for col in range(1, scan_cols + 1):
            text = _cell_str(ws, row, col).casefold()
            if start_col is None and ("region" in text or text == "regions"):
                start_col = col
            if any(token in text for token in ("budget", "overtime", "total hours", "total hour")):
                end_col = max(end_col, col)
            if text:
                end_col = max(end_col, col)

        if start_col is None:
            continue

        # Expand end_col across merged/spanned metric columns under TOTAL HOURS.
        # Typical layout uses ~6–7 value cols per block + spacer → ~20–22 cols total.
        for col in range(end_col + 1, min(start_col + 24, scan_cols + 1)):
            # Keep extending while this row or the next two header rows have content.
            has = any(_cell_str(ws, r, col) for r in (row, row + 1, row + 2))
            if has:
                end_col = col
            elif col - end_col > 1:
                break

        # Ensure we cover the third header row's last MTD/YTD cells.
        for col in range(end_col, min(start_col + 24, scan_cols + 1)):
            if any(_cell_str(ws, r, col) for r in (row, row + 1, row + 2)):
                end_col = col

        return row, start_col, end_col
    return None


def _discover_finance_table(workbook_bytes: bytes, sheet_name: str) -> tuple[int, int, int, int]:
    """Locate the Regions BUDGET / OVERTIME / TOTAL HOURS table (headers + data)."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    try:
        match = _find_sheet_name(list(wb.sheetnames), sheet_name) or sheet_name
        ws = wb[match]
        hdr = _find_finance_header(ws)
        if hdr is None:
            raise ValueError(
                f"Could not find Regions/BUDGET/OVERTIME/TOTAL HOURS header on {sheet_name!r}"
            )

        start_row, start_col, end_col = hdr
        # Include MTD/YTD/Score and Actual/(B)/W Goal sub-header rows when present.
        header_rows = 1
        for extra in (1, 2):
            blob = _row_blob(ws, start_row + extra, end_col)
            if any(token in blob for token in ("mtd", "ytd", "score", "actual", "goal")):
                header_rows = extra + 1
            else:
                break

        end_row = start_row + header_rows - 1
        max_scan = min(int(ws.max_row or 80), start_row + 40)
        blank_streak = 0
        for row in range(start_row + header_rows, max_scan + 1):
            label = _cell_str(ws, row, start_col).casefold()
            row_has = any(_cell_str(ws, row, col) for col in range(start_col, end_col + 1))
            if not row_has:
                blank_streak += 1
                if blank_streak >= 2:
                    break
                continue
            blank_streak = 0
            if any(token in label for token in _END_TOKENS):
                break
            if label and label in {"finance", "kpi", "yr_nb", "entity"}:
                break
            end_row = row

        if end_row < start_row + header_rows:
            end_row = min(max_scan, start_row + header_rows + 8)

        logger.info(
            "FINANCE Regions table %s!%s",
            match,
            _range_address(start_row, end_row, start_col, end_col),
        )
        print(
            f">>> FINANCE table range: {match}!"
            f"{_range_address(start_row, end_row, start_col, end_col)} "
            f"(Regions / BUDGET / OVERTIME / TOTAL HOURS)"
        )
        return start_row, end_row, start_col, end_col
    finally:
        wb.close()


def _remove_finance_content_pictures(slide) -> int:
    """Remove the template finance overview picture (keep title / footer text)."""
    removed = 0
    left, top, width, height = FINANCE_TABLE_BOX
    band_bottom = top + height + 250_000
    for shape in list(slide.shapes):
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        if int(shape.top) >= band_bottom:
            continue
        if int(shape.top) + int(shape.height) < top - 100_000:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
    return removed


def apply_finance_workings_panels(slide, data, element: dict) -> bool:
    """Fill slide 8 from Workings!FINANCE Regions/Budget/OT/Hours table screenshot."""
    workbook = element.get("workbook", "workings")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "fill")).lower()

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("FINANCE workings screenshot skipped: %s", exc)
        return False

    available = []
    try:
        available = list(data.sheet_names(workbook))
    except Exception:
        available = []

    try:
        sheet_name = resolve_sheet_name(
            workbook_bytes,
            sheet=element.get("sheet", "FINANCE"),
            sheet_index=element.get("sheet_index"),
            sheet_match=element.get("sheet_match", ["FINANCE", "Finance"]),
            sheet_match_index=int(element.get("sheet_match_index", 0) or 0),
            available=available or None,
        )
    except Exception as exc:
        logger.warning("Could not resolve Workings FINANCE sheet: %s", exc)
        return False

    try:
        start_row, end_row, start_col, end_col = _discover_finance_table(workbook_bytes, sheet_name)
        png = capture_range_png(
            workbook_bytes,
            sheet_name,
            start_row,
            end_row,
            start_col,
            end_col,
            prefer_excel_com=prefer_com,
        )
        png = _validate_capture(
            png,
            label=f"FINANCE table {sheet_name}",
            min_w=80,
            min_h=40,
            require_wide=False,
        )
    except Exception as exc:
        logger.warning("FINANCE table capture failed: %s", exc)
        print(f">>> ERROR: FINANCE table screenshot failed: {exc}")
        return False

    out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    removed = _remove_finance_content_pictures(slide)
    logger.info("Slide 8 cleared %s picture(s) before FINANCE screenshot", removed)

    left, top, width, height = FINANCE_TABLE_BOX
    place_picture_on_slide(
        slide,
        png,
        left=left,
        top=top,
        max_width=width,
        max_height=height,
        fit=fit,
    )

    try:
        with Image.open(io.BytesIO(png)) as img:
            debug = out_dir / f"_debug_finance_table_{img.width}x{img.height}.png"
            img.save(debug)
            print(
                f">>> FINANCE Regions/BUDGET/OVERTIME/TOTAL HOURS screenshot placed "
                f"from workings/{sheet_name} ({img.width}x{img.height}) -> {debug.name}"
            )
    except Exception:
        print(f">>> FINANCE table screenshot placed from workings/{sheet_name}")

    print(
        f"\n>>> Slide 8 Finance: Regions/BUDGET/OVERTIME/TOTAL HOURS screenshot "
        f"from workings/{sheet_name}\n"
    )
    return True
