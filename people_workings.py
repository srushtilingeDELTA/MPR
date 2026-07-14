"""People slide (PPT 7) from New GSE MPR Workings.xlsx → PEOPLE tab.

Screenshots only:
  1. The first PEOPLE MTD/YTD/Score table
  2. The three Excel graphs on that tab:
       - Leadership Engagement
       - Psychological Safety
       - Accountability

Charts are captured with Excel COM CopyPicture / Chart.Export (true screenshots),
then placed in a right-side stack. No native PowerPoint chart rebuilds.
"""

from __future__ import annotations

import io
import logging
import tempfile
import time
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from gir_panels import clear_leading_action_narrative
from scorecard_screenshots import (
    _clipboard_image,
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

# Right-side chart column — three equal stacked slots.
PEOPLE_CHART_LEFT = 7_976_668
PEOPLE_CHART_WIDTH = 3_657_600
PEOPLE_CHART_TOP = 1_036_821
PEOPLE_CHART_BOTTOM = 6_200_000
PEOPLE_CHART_GAP = 70_000

# Required charts, in top→bottom display order.
PEOPLE_CHART_SPECS = [
    {
        "key": "leadership",
        "title": "Leadership Engagement",
        "match": ["leadership engagement", "leadership"],
    },
    {
        "key": "psychological",
        "title": "Psychological Safety",
        "match": ["psychological safety", "psychological", "psych safety"],
    },
    {
        "key": "accountability",
        "title": "Accountability",
        "match": ["accountability"],
    },
]

_TABLE_END_TOKENS = (
    "where i work, employees are held accountable",
    "employees are held accountable",
)


def people_chart_boxes(count: int = 3) -> list[tuple[int, int, int, int]]:
    """Evenly stack `count` chart boxes in the right column."""
    count = max(1, int(count))
    usable = PEOPLE_CHART_BOTTOM - PEOPLE_CHART_TOP - PEOPLE_CHART_GAP * (count - 1)
    height = max(900_000, usable // count)
    boxes = []
    top = PEOPLE_CHART_TOP
    for idx in range(count):
        h = height
        if idx == count - 1:
            h = max(height, PEOPLE_CHART_BOTTOM - top)
        boxes.append((PEOPLE_CHART_LEFT, top, PEOPLE_CHART_WIDTH, h))
        top += height + PEOPLE_CHART_GAP
    return boxes


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
            row_blob = " ".join(_cell_str(ws, row, c).casefold() for c in range(col, min(col + 8, max_col + 1)))
            next_blob = " ".join(
                _cell_str(ws, row + 1, c).casefold() for c in range(col, min(col + 8, max_col + 1))
            )
            if "mtd" in row_blob or "ytd" in row_blob or "score" in row_blob:
                return row, col
            if "actual" in next_blob or "goal" in next_blob:
                return row, col
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
            if row > start_row + 2 and label in {"people", "kpi", "yr_nb", "entity"}:
                end_row = row - 1
                break

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


def _com_chart_title(chart_obj) -> str:
    for getter in (
        lambda: chart_obj.Chart.ChartTitle.Text,
        lambda: chart_obj.Chart.HasTitle and chart_obj.Chart.ChartTitle.Text,
        lambda: chart_obj.Name,
    ):
        try:
            text = str(getter() or "").strip()
            if text:
                return text
        except Exception:
            continue
    return ""


def _match_chart_spec(title: str) -> dict | None:
    lower = (title or "").casefold()
    for spec in PEOPLE_CHART_SPECS:
        for token in spec["match"]:
            if token in lower:
                return spec
    return None


def _screenshot_excel_chart(chart_obj, *, label: str) -> bytes:
    """True screenshot of an Excel chart object (CopyPicture, then Chart.Export)."""
    errors: list[str] = []

    # 1) CopyPicture of the chart object / chart area → clipboard PNG.
    for appearance, fmt in ((1, 2), (2, 2)):  # xlScreen/xlPrinter + xlBitmap
        for copier in (
            lambda: chart_obj.CopyPicture(Appearance=appearance, Format=fmt),
            lambda: chart_obj.Chart.CopyPicture(Appearance=appearance, Format=fmt),
            lambda: chart_obj.Chart.ChartArea.Copy(),
        ):
            try:
                try:
                    import win32clipboard  # type: ignore

                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
                copier()
                for _ in range(40):
                    time.sleep(0.15)
                    grabbed = _clipboard_image()
                    if grabbed is None:
                        continue
                    png = _png_from_pil(grabbed)
                    return _validate_capture(
                        png, label=label, min_w=80, min_h=60, require_wide=False
                    )
            except Exception as exc:
                errors.append(f"CopyPicture:{exc}")

    # 2) Chart.Export fallback.
    export_path = Path(tempfile.gettempdir()) / f"mpr_people_chart_{time.time_ns()}.png"
    try:
        if export_path.exists():
            export_path.unlink()
        chart_obj.Chart.Export(str(export_path))
        if export_path.exists() and export_path.stat().st_size > 1500:
            data = export_path.read_bytes()
            try:
                return _validate_capture(
                    data, label=f"{label} Export", min_w=80, min_h=60, require_wide=False
                )
            except Exception:
                with Image.open(io.BytesIO(data)) as img:
                    return _png_from_pil(img)
    except Exception as exc:
        errors.append(f"Export:{exc}")
    finally:
        try:
            if export_path.exists():
                export_path.unlink()
        except Exception:
            pass

    raise RuntimeError(f"{label} screenshot failed ({'; '.join(errors) or 'no image'})")


def _screenshot_named_people_charts(workbook_path: Path, sheet_name: str) -> dict[str, bytes]:
    """Screenshot PEOPLE sheet charts keyed by leadership/psychological/accountability."""
    excel = None
    wb = None
    by_key: dict[str, bytes] = {}
    ordered_fallback: list[bytes] = []
    try:
        excel, wb = _open_excel_workbook(workbook_path)
        ws = _find_com_worksheet(wb, sheet_name)
        ws.Activate()
        try:
            excel.Goto(ws.Range("A1"), True)
        except Exception:
            pass
        time.sleep(0.4)

        count = int(ws.ChartObjects().Count)
        print(f">>> PEOPLE sheet has {count} Excel chart object(s) to screenshot")
        if count < 3:
            logger.warning("Expected 3 PEOPLE charts, Excel reports %s ChartObjects", count)

        scored: list[tuple[float, str, bytes]] = []
        for idx in range(1, count + 1):
            chart_obj = ws.ChartObjects(idx)
            title = _com_chart_title(chart_obj)
            label = f"PEOPLE chart {idx} ({title or 'untitled'})"
            try:
                # Bring chart into view so CopyPicture is not blank.
                try:
                    chart_obj.Activate()
                except Exception:
                    pass
                time.sleep(0.2)
                top = float(getattr(chart_obj, "Top", idx * 100))
                data = _screenshot_excel_chart(chart_obj, label=label)
                scored.append((top, title, data))
                print(f">>> Screenshotted Excel graph {idx}/{count}: {title!r} ({len(data)} bytes)")
            except Exception as exc:
                logger.warning("%s failed: %s", label, exc)
                print(f">>> WARNING: could not screenshot PEOPLE chart {idx}: {exc}")

        scored.sort(key=lambda item: item[0])
        for top, title, data in scored:
            spec = _match_chart_spec(title)
            if spec and spec["key"] not in by_key:
                by_key[spec["key"]] = data
                print(f">>> Matched screenshot to '{spec['title']}' (Excel title {title!r})")
            else:
                ordered_fallback.append(data)

        # If titles are blank/generic, assign remaining screenshots top→bottom.
        for spec in PEOPLE_CHART_SPECS:
            if spec["key"] in by_key:
                continue
            if not ordered_fallback:
                break
            by_key[spec["key"]] = ordered_fallback.pop(0)
            print(f">>> Assigned position-ordered screenshot to '{spec['title']}'")
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
    return by_key


def _remove_people_data_shapes(slide) -> int:
    """Remove native PEOPLE table/charts and prior content pictures (keep logo)."""
    removed = 0
    for shape in list(slide.shapes):
        drop = False
        if getattr(shape, "has_table", False):
            drop = True
        elif getattr(shape, "has_chart", False):
            drop = True
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE and 900_000 <= int(shape.top) < 6_000_000:
            drop = True
        if not drop:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
    return removed


def apply_people_workings_panels(slide, data, element: dict) -> bool:
    """Fill slide 7 from Workings!PEOPLE screenshots: first table + 3 Excel graphs."""
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

    # 1) First PEOPLE metrics table screenshot.
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

    # 2) Screenshot the three Excel graphs from the PEOPLE tab (COM required).
    charts_by_key: dict[str, bytes] = {}
    if not prefer_com:
        print(">>> ERROR: PEOPLE charts require Excel COM screenshots (prefer_excel_com=true)")
    else:
        try:
            with tempfile.TemporaryDirectory(prefix="mpr_people_") as tmp:
                path = Path(tmp) / "workings.xlsx"
                path.write_bytes(workbook_bytes)
                charts_by_key = _screenshot_named_people_charts(path, sheet_name)
        except Exception as exc:
            logger.error("PEOPLE Excel chart screenshots unavailable: %s", exc)
            print(f">>> ERROR: could not screenshot PEOPLE graphs from Excel: {exc}")

    if not table_png and not charts_by_key:
        logger.warning("No PEOPLE table/chart screenshots from workings/%s", sheet_name)
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
                    f">>> PEOPLE table screenshot placed from workings/{sheet_name} "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PEOPLE table screenshot placed from workings/{sheet_name}")

    boxes = people_chart_boxes(chart_limit)
    specs = PEOPLE_CHART_SPECS[:chart_limit]
    for idx, (spec, box) in enumerate(zip(specs, boxes)):
        png = charts_by_key.get(spec["key"])
        if png is None:
            print(f">>> WARNING: missing Excel screenshot for '{spec['title']}'")
            continue
        left, top, width, height = box
        place_picture_on_slide(
            slide,
            png,
            left=left,
            top=top,
            max_width=width,
            max_height=height,
            fit=fit,
        )
        placed += 1
        try:
            with Image.open(io.BytesIO(png)) as img:
                debug = out_dir / f"_debug_people_{spec['key']}_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> PEOPLE graph {idx+1}/3 '{spec['title']}' screenshot placed "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PEOPLE graph {idx+1}/3 '{spec['title']}' screenshot placed")

    missing = [spec["title"] for spec in specs if spec["key"] not in charts_by_key]
    if missing:
        logger.warning("Missing PEOPLE Excel graph screenshots: %s", ", ".join(missing))
        print(
            ">>> WARNING: these PEOPLE graphs were NOT screenshotted from Excel: "
            + ", ".join(missing)
        )
    else:
        print(
            ">>> PEOPLE right-side Excel graph screenshots: "
            "Leadership Engagement, Psychological Safety, Accountability"
        )

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> PEOPLE Leading Issues / Action Plans cleared ({n} text box(es))")

    print(
        f"\n>>> Slide 7 People: placed {placed} screenshot(s) from workings/{sheet_name} "
        f"(table={'yes' if table_png else 'no'}, "
        f"graph screenshots={len(charts_by_key)}/3)\n"
    )
    return placed > 0
