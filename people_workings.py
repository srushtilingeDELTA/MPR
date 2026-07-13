"""People slide (PPT 7) from New GSE MPR Workings.xlsx → PEOPLE tab.

Only the first PEOPLE metrics table matters, plus the three Excel charts on
that sheet. Screenshots go into the template table slot and three right-side
chart slots. Leading Issues / Action Plans stay empty editable text boxes.
"""

from __future__ import annotations

import io
import logging
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from gir_panels import clear_leading_action_narrative
from scorecard_screenshots import (
    _find_com_worksheet,
    _find_sheet_name,
    _open_excel_workbook,
    _png_from_pil,
    _range_address,
    _validate_capture,
    capture_range_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Template table slot from GSE MPR template slide 7.
PEOPLE_TABLE_BOX = (362_427, 1_036_821, 7_426_936, 2_255_520)

# Three chart slots stacked on the right (top slot added above the two template charts).
PEOPLE_CHART_BOXES = [
    (7_976_668, 1_036_821, 3_657_600, 1_700_000),  # top — beside table
    (7_976_668, 2_817_996, 3_657_600, 1_800_225),  # middle — template Chart 28
    (7_976_668, 4_675_371, 3_657_600, 1_792_605),  # bottom — template Chart 29
]

# Hard stop markers once the first PEOPLE scorecard table is complete.
_TABLE_END_TOKENS = (
    "where i work, employees are held accountable",
    "employees are held accountable",
)


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row, col).value
    if val is None:
        return ""
    return str(val).strip()


def _find_people_header(ws, *, max_row: int = 40, max_col: int = 12) -> tuple[int, int] | None:
    """Find the first PEOPLE scorecard header (PEOPLE + MTD/YTD/Score)."""
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            text = _cell_str(ws, row, col).casefold()
            if text != "people":
                continue
            # Confirm this is the scorecard header row, not a random label.
            row_blob = " ".join(_cell_str(ws, row, c).casefold() for c in range(col, min(col + 8, max_col + 1)))
            next_blob = " ".join(
                _cell_str(ws, row + 1, c).casefold() for c in range(col, min(col + 8, max_col + 1))
            )
            if "mtd" in row_blob or "ytd" in row_blob or "score" in row_blob:
                return row, col
            if "actual" in next_blob or "goal" in next_blob:
                return row, col
    # Fallback: first Leadership Engagement block.
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            if "leadership engagement" in _cell_str(ws, row, col).casefold():
                return max(1, row - 2), col
    return None


def _discover_first_people_table(workbook_bytes: bytes, sheet_name: str) -> tuple[int, int, int, int]:
    """Locate only the first PEOPLE MTD/YTD/Score table (ignore later tables/dumps)."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    try:
        match = _find_sheet_name(list(wb.sheetnames), sheet_name) or sheet_name
        ws = wb[match]
        hdr = _find_people_header(ws)
        if hdr is None:
            raise ValueError(f"Could not find PEOPLE scorecard header on {sheet_name!r}")

        start_row, start_col = hdr
        # Typical table is 7 columns: label + MTD Actual/B(W) + YTD Actual/B(W) + Score MTD/YTD.
        end_col = start_col + 6

        end_row = start_row
        for row in range(start_row, min(int(ws.max_row or 40), start_row + 16) + 1):
            label = _cell_str(ws, row, start_col).casefold()
            row_has = any(_cell_str(ws, row, col) for col in range(start_col, end_col + 1))
            if not row_has:
                break
            end_row = row
            if any(token in label for token in _TABLE_END_TOKENS):
                break
            # Stop before a second major section / raw dump.
            if row > start_row + 2 and label in {"people", "kpi", "yr_nb", "entity"}:
                end_row = row - 1
                break

        # Ensure we captured through Accountability when present.
        if end_row < start_row + 8:
            end_row = min(int(ws.max_row or 40), start_row + 9)

        logger.info(
            "PEOPLE first table %s!%s",
            match,
            _range_address(start_row, end_row, start_col, end_col),
        )
        return start_row, end_row, start_col, end_col
    finally:
        wb.close()


def _export_sheet_charts(workbook_path: Path, sheet_name: str, *, limit: int = 3) -> list[bytes]:
    """Export embedded Excel charts on a sheet as PNG bytes (top-to-bottom order)."""
    excel = None
    wb = None
    pngs: list[tuple[float, bytes]] = []
    try:
        excel, wb = _open_excel_workbook(workbook_path)
        ws = _find_com_worksheet(wb, sheet_name)
        ws.Activate()
        count = int(ws.ChartObjects().Count)
        for idx in range(1, count + 1):
            chart_obj = ws.ChartObjects(idx)
            export_path = Path(tempfile.gettempdir()) / f"mpr_people_chart_{idx}.png"
            try:
                if export_path.exists():
                    export_path.unlink()
                top = float(getattr(chart_obj, "Top", idx * 100))
                chart_obj.Chart.Export(str(export_path))
                if export_path.exists() and export_path.stat().st_size > 1500:
                    data = export_path.read_bytes()
                    try:
                        data = _validate_capture(
                            data, label=f"PEOPLE chart {idx}", min_w=80, min_h=60, require_wide=False
                        )
                    except Exception:
                        with Image.open(io.BytesIO(data)) as img:
                            data = _png_from_pil(img)
                    pngs.append((top, data))
                    logger.info("Exported PEOPLE Excel chart %s (%s bytes)", idx, len(data))
            except Exception as exc:
                logger.info("PEOPLE chart %s export failed: %s", idx, exc)
            finally:
                try:
                    if export_path.exists():
                        export_path.unlink()
                except Exception:
                    pass
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
    pngs.sort(key=lambda item: item[0])
    return [png for _, png in pngs[:limit]]


def _remove_people_data_shapes(slide) -> int:
    """Remove native PEOPLE table/charts in the content area (keep logo + narrative chrome)."""
    removed = 0
    for shape in list(slide.shapes):
        drop = False
        if getattr(shape, "has_table", False):
            drop = True
        elif getattr(shape, "has_chart", False):
            drop = True
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            top = int(shape.top)
            if 900_000 <= top < 6_000_000:
                drop = True
        if not drop:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
    return removed


def apply_people_workings_panels(slide, data, element: dict) -> bool:
    """Fill slide 7 from Workings!PEOPLE: first table + up to three charts."""
    workbook = element.get("workbook", "workings")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "fill")).lower()
    chart_limit = int(element.get("chart_count", 3) or 3)

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("PEOPLE workings screenshot skipped: %s", exc)
        return False

    available = []
    try:
        available = list(data.sheet_names(workbook))
    except Exception:
        available = []

    try:
        sheet_name = resolve_sheet_name(
            workbook_bytes,
            sheet=element.get("sheet", "PEOPLE"),
            sheet_index=element.get("sheet_index"),
            sheet_match=element.get("sheet_match", ["PEOPLE", "People"]),
            sheet_match_index=int(element.get("sheet_match_index", 0) or 0),
            available=available or None,
        )
    except Exception as exc:
        logger.warning("Could not resolve Workings PEOPLE sheet: %s", exc)
        return False

    out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    placed = 0

    # 1) First PEOPLE metrics table only.
    table_png = None
    try:
        start_row, end_row, start_col, end_col = _discover_first_people_table(workbook_bytes, sheet_name)
        table_png = capture_range_png(
            workbook_bytes,
            sheet_name,
            start_row,
            end_row,
            start_col,
            end_col,
            prefer_excel_com=prefer_com,
        )
        table_png = _validate_capture(
            table_png, label=f"PEOPLE table {sheet_name}", min_w=80, min_h=40, require_wide=False
        )
    except Exception as exc:
        logger.warning("PEOPLE first-table capture failed: %s", exc)
        table_png = None

    # 2) Up to three Excel charts from the PEOPLE sheet.
    chart_pngs: list[bytes] = []
    if prefer_com:
        try:
            with tempfile.TemporaryDirectory(prefix="mpr_people_") as tmp:
                path = Path(tmp) / "workings.xlsx"
                path.write_bytes(workbook_bytes)
                chart_pngs = _export_sheet_charts(path, sheet_name, limit=chart_limit)
        except Exception as exc:
            logger.info("PEOPLE Excel chart export unavailable: %s", exc)

    if not table_png and not chart_pngs:
        logger.warning("No PEOPLE table/charts captured from workings/%s", sheet_name)
        return False

    removed = _remove_people_data_shapes(slide)
    logger.info("Slide 7 cleared %s table/chart/picture shape(s)", removed)

    if table_png:
        left, top, width, height = PEOPLE_TABLE_BOX
        place_picture_on_slide(
            slide,
            table_png,
            left=left,
            top=top,
            max_width=width,
            max_height=height,
            fit=fit,
        )
        placed += 1
        try:
            with Image.open(io.BytesIO(table_png)) as img:
                debug = out_dir / f"_debug_people_table_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> PEOPLE first table placed from workings/{sheet_name} "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PEOPLE first table placed from workings/{sheet_name}")

    boxes = PEOPLE_CHART_BOXES[:chart_limit]
    for idx, box in enumerate(boxes):
        if idx >= len(chart_pngs):
            break
        left, top, width, height = box
        place_picture_on_slide(
            slide,
            chart_pngs[idx],
            left=left,
            top=top,
            max_width=width,
            max_height=height,
            fit=fit,
        )
        placed += 1
        try:
            with Image.open(io.BytesIO(chart_pngs[idx])) as img:
                debug = out_dir / f"_debug_people_chart{idx+1}_{img.width}x{img.height}.png"
                img.save(debug)
                print(f">>> PEOPLE chart {idx+1}/3 placed ({img.width}x{img.height}) -> {debug.name}")
        except Exception:
            print(f">>> PEOPLE chart {idx+1}/3 placed")

    if chart_pngs and len(chart_pngs) < 3:
        logger.warning("PEOPLE sheet exported %s chart(s); expected 3", len(chart_pngs))
        print(f">>> WARNING: expected 3 PEOPLE charts, exported {len(chart_pngs)}")
    elif not chart_pngs:
        print(">>> WARNING: no PEOPLE charts exported (Excel COM required on Windows)")

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> PEOPLE Leading Issues / Action Plans cleared ({n} text box(es))")

    print(
        f"\n>>> Slide 7 People: placed {placed} screenshot(s) from workings/{sheet_name} "
        f"(first table={'yes' if table_png else 'no'}, charts={min(len(chart_pngs), len(boxes))}/3)\n"
    )
    return placed > 0
