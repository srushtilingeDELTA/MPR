"""People slide (PPT 7) from New GSE MPR Workings.xlsx → PEOPLE tab.

Screenshots the PEOPLE dashboard table into the template table slot and
exports Excel charts into the two right-side chart slots. Leading Issues /
Action Plans stay empty editable text boxes.
"""

from __future__ import annotations

import io
import logging
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from gir_panels import clear_leading_action_narrative
from scorecard_screenshots import (
    _find_com_worksheet,
    _find_sheet_name,
    _open_excel_workbook,
    _png_from_pil,
    _range_address,
    _used_bounds,
    _validate_capture,
    capture_range_png,
    capture_sheet_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Template slots from GSE MPR - Template.pptx slide 7 (index 6).
PEOPLE_TABLE_BOX = (362_427, 1_036_821, 7_426_936, 2_255_520)
PEOPLE_CHART_BOXES = [
    (7_976_668, 2_817_996, 3_657_600, 1_800_225),  # Chart 28
    (7_976_668, 4_675_371, 3_657_600, 1_792_605),  # Chart 29
]


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row, col).value
    if val is None:
        return ""
    return str(val).strip()


def _find_token(ws, tokens: list[str], *, max_row: int = 60, max_col: int = 20) -> tuple[int, int] | None:
    needles = [t.casefold() for t in tokens]
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            text = _cell_str(ws, row, col).casefold()
            if not text:
                continue
            for needle in needles:
                if needle == text or needle in text:
                    return row, col
    return None


def _dashboard_end_col(ws) -> int:
    markers = ("yr_nb", "all", "mo_nb", "kpi", "entity", "num", "den")
    for row in range(1, 30):
        for col in range(8, min(40, int(ws.max_column or 40)) + 1):
            text = _cell_str(ws, row, col).casefold()
            if text in markers:
                return max(7, col - 1)
    return min(14, int(ws.max_column or 14))


def _discover_people_table_range(workbook_bytes: bytes, sheet_name: str) -> tuple[int, int, int, int]:
    """Return 1-based (start_row, end_row, start_col, end_col) for the PEOPLE metrics table."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    try:
        match = _find_sheet_name(list(wb.sheetnames), sheet_name) or sheet_name
        ws = wb[match]
        max_col = _dashboard_end_col(ws)
        hdr = _find_token(
            ws,
            ["people", "leadership engagement", "psychological safety", "accountability"],
            max_col=max_col,
        )
        if hdr is None:
            min_row, max_row, min_col, max_c = _used_bounds(ws)
            return min_row, min(max_row, min_row + 25), min_col, min(max_c, max_col)

        start_row, start_col = hdr
        # Include a title/header row above when present.
        if start_row > 1:
            above = _cell_str(ws, start_row - 1, start_col).casefold()
            if above and ("people" in above or "mtd" in above or "actual" in above):
                start_row -= 1

        end_row = start_row
        empty_streak = 0
        for row in range(start_row, min(int(ws.max_row or 80), start_row + 35) + 1):
            row_has = any(_cell_str(ws, row, col) for col in range(start_col, max_col + 1))
            if row_has:
                end_row = row
                empty_streak = 0
            else:
                empty_streak += 1
                if empty_streak >= 2 and end_row > start_row:
                    break

        end_col = start_col
        for col in range(start_col, max_col + 1):
            if any(_cell_str(ws, row, col) for row in range(start_row, end_row + 1)):
                end_col = col

        # Prefer at least the typical PEOPLE scorecard width.
        end_col = max(end_col, min(max_col, start_col + 6))
        logger.info(
            "PEOPLE table range %s!%s",
            match,
            _range_address(start_row, end_row, start_col, end_col),
        )
        return start_row, end_row, start_col, end_col
    finally:
        wb.close()


def _export_sheet_charts(workbook_path: Path, sheet_name: str) -> list[bytes]:
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
    return [png for _, png in pngs]


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
            # Keep footer logo.
            if top < 6_000_000:
                # Only remove pictures that sit in the table/chart band.
                if top >= 900_000:
                    drop = True
        if not drop:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
    return removed


def apply_people_workings_panels(slide, data, element: dict) -> bool:
    """Fill slide 7 from Workings!PEOPLE screenshots (table + charts)."""
    workbook = element.get("workbook", "workings")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "fill")).lower()

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

    # 1) PEOPLE metrics table screenshot.
    table_png = None
    try:
        start_row, end_row, start_col, end_col = _discover_people_table_range(workbook_bytes, sheet_name)
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
        logger.warning("PEOPLE table capture failed (%s); trying full sheet", exc)
        try:
            table_png = capture_sheet_png(
                workbook_bytes,
                sheet_name,
                prefer_excel_com=prefer_com,
                max_rows=element.get("max_rows", 40),
                max_cols=element.get("max_cols", 12),
            )
            table_png = _validate_capture(
                table_png, label=f"PEOPLE sheet {sheet_name}", min_w=80, min_h=40, require_wide=False
            )
        except Exception as exc2:
            logger.warning("PEOPLE full-sheet capture also failed: %s", exc2)
            table_png = None

    # 2) Excel chart exports (preferred for the two right-side slots).
    chart_pngs: list[bytes] = []
    if prefer_com:
        try:
            with tempfile.TemporaryDirectory(prefix="mpr_people_") as tmp:
                path = Path(tmp) / "workings.xlsx"
                path.write_bytes(workbook_bytes)
                chart_pngs = _export_sheet_charts(path, sheet_name)
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
                    f">>> PEOPLE table placed from workings/{sheet_name} "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PEOPLE table placed from workings/{sheet_name}")

    for idx, box in enumerate(PEOPLE_CHART_BOXES):
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
                print(f">>> PEOPLE chart {idx+1} placed ({img.width}x{img.height}) -> {debug.name}")
        except Exception:
            print(f">>> PEOPLE chart {idx+1} placed")

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> PEOPLE Leading Issues / Action Plans cleared ({n} text box(es))")

    print(
        f"\n>>> Slide 7 People: placed {placed} screenshot(s) from workings/{sheet_name} "
        f"(table={'yes' if table_png else 'no'}, charts={min(len(chart_pngs), len(PEOPLE_CHART_BOXES))})\n"
    )
    return placed > 0
