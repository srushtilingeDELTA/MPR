"""PMI Compliance slide (PPT 11) from New GSE MPR Workings.xlsx → PMI tab.

Screenshots:
  1. Regions tables (Motorized/Stationary + Non-Motorized) into the template OLE slot
  2. Motorized and Stationary Excel graphs into the bottom chart slots

Leading Issues / Action Plans are cleared to empty editable text boxes.
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
from people_workings import (
    _com_chart_title,
    _screenshot_excel_chart,
)
from scorecard_screenshots import (
    _find_com_worksheet,
    _find_sheet_name,
    _open_excel_workbook,
    _range_address,
    _validate_capture,
    capture_range_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Template OLE / table slot (Object 3 on PMI COMPLIANCE slide).
PMI_TABLE_BOX = (314_628, 1_073_741, 7_994_650, 3_937_001)

# Bottom native chart slots.
PMI_CHART_BOXES = {
    "motorized": (1_331_718, 5_353_707, 4_392_295, 1_353_312),
    "stationary": (5_706_868, 5_289_350, 4_389_120, 1_355_345),
}

PMI_CHART_SPECS = [
    {"key": "motorized", "title": "Motorized", "match": ["motorized", "motorised", "pm (m)"]},
    {"key": "stationary", "title": "Stationary", "match": ["stationary", "pm (s)"]},
]

_TABLE_END_TOKENS = (
    "month",
    "jan",
    "leading issue",
    "action plan",
    "notes",
    "motorized actual",
    "stationary actual",
)


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row, col).value
    if val is None:
        return ""
    return str(val).strip()


def _row_blob(ws, row: int, max_col: int) -> str:
    return " ".join(_cell_str(ws, row, c).casefold() for c in range(1, max_col + 1))


def _find_pmi_header(ws, *, max_row: int = 40, max_col: int = 40) -> tuple[int, int] | None:
    """Find the PMI Regions / MTD / YTD / Score header."""
    scan_cols = min(max(int(ws.max_column or max_col), 20), max_col)
    for row in range(1, min(int(ws.max_row or max_row), max_row) + 1):
        blob = _row_blob(ws, row, scan_cols)
        if "region" not in blob:
            continue
        if "mtd" not in blob and "ytd" not in blob and "score" not in blob:
            # Sub-headers may be on the next row.
            next_blob = _row_blob(ws, row + 1, scan_cols)
            if "mtd" not in next_blob and "actual" not in next_blob and "score" not in next_blob:
                continue
        for col in range(1, scan_cols + 1):
            text = _cell_str(ws, row, col).casefold()
            if "region" in text:
                return row, col
    return None


def _discover_pmi_table(workbook_bytes: bytes, sheet_name: str) -> tuple[int, int, int, int]:
    """Locate the PMI Regions tables (Motorized/Stationary + Non-Motorized)."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    try:
        match = _find_sheet_name(list(wb.sheetnames), sheet_name) or sheet_name
        ws = wb[match]
        hdr = _find_pmi_header(ws)
        if hdr is None:
            raise ValueError(f"Could not find PMI Regions/MTD/YTD header on {sheet_name!r}")

        start_row, start_col = hdr
        scan_cols = min(max(int(ws.max_column or 30), 20), 40)

        # Include Non-Motorized columns to the right of the first score block.
        end_col = start_col
        for row in (start_row, start_row + 1, start_row + 2):
            for col in range(start_col, scan_cols + 1):
                text = _cell_str(ws, row, col).casefold()
                if text or any(
                    token in text
                    for token in ("mtd", "ytd", "score", "actual", "goal", "non-motor", "non motor")
                ):
                    if _cell_str(ws, row, col):
                        end_col = max(end_col, col)

        # Prefer extending through a NON-MOTORIZED title if present nearby.
        for row in range(max(1, start_row - 2), start_row + 4):
            for col in range(start_col, scan_cols + 1):
                text = _cell_str(ws, row, col).casefold()
                if "non-motor" in text or "non motor" in text or text == "nme":
                    end_col = max(end_col, col + 6)

        # Grow while header/data rows have content (don't stop at one spacer).
        for col in range(end_col, scan_cols + 1):
            if any(_cell_str(ws, r, col) for r in range(start_row, start_row + 4)):
                end_col = col

        header_rows = 1
        for extra in (1, 2):
            blob = _row_blob(ws, start_row + extra, end_col)
            if any(token in blob for token in ("mtd", "ytd", "score", "actual", "goal")):
                header_rows = extra + 1
            else:
                break

        end_row = start_row + header_rows - 1
        max_scan = min(int(ws.max_row or 80), start_row + 50)
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
            if any(token in label for token in _TABLE_END_TOKENS):
                break
            if label in {"pmi", "kpi", "yr_nb", "entity", "month"}:
                break
            end_row = row

        if end_row < start_row + header_rows:
            end_row = min(max_scan, start_row + header_rows + 12)

        addr = _range_address(start_row, end_row, start_col, end_col)
        logger.info("PMI table %s!%s", match, addr)
        print(f">>> PMI table range: {match}!{addr}")
        return start_row, end_row, start_col, end_col
    finally:
        wb.close()


def _match_pmi_chart(title: str) -> dict | None:
    lower = (title or "").casefold()
    for spec in PMI_CHART_SPECS:
        for token in spec["match"]:
            if token in lower:
                return spec
    return None


def _iter_sheet_charts(ws) -> list[tuple[float, float, str, object]]:
    """Collect Excel charts on a sheet as (top, left, title, obj)."""
    found: list[tuple[float, float, str, object]] = []
    seen: set[str] = set()

    def _add(obj, *, source: str) -> None:
        try:
            name = str(getattr(obj, "Name", "") or "")
        except Exception:
            name = ""
        key = name.casefold() or f"{source}:{id(obj)}"
        if key in seen:
            return
        seen.add(key)
        try:
            top = float(getattr(obj, "Top", len(found) * 100))
            left = float(getattr(obj, "Left", len(found) * 100))
        except Exception:
            top, left = float(len(found) * 100), 0.0
        title = _com_chart_title(obj) or name
        found.append((top, left, title, obj))
        print(f">>> Found PMI Excel graph ({source}): {title or 'untitled'!r}")

    try:
        count = int(ws.ChartObjects().Count)
    except Exception:
        count = 0
    for idx in range(1, count + 1):
        try:
            _add(ws.ChartObjects(idx), source=f"ChartObjects[{idx}]")
        except Exception as exc:
            logger.warning("PMI ChartObjects(%s) unreadable: %s", idx, exc)

    try:
        shape_count = int(ws.Shapes.Count)
    except Exception:
        shape_count = 0
    for idx in range(1, shape_count + 1):
        try:
            shape = ws.Shapes(idx)
            if int(getattr(shape, "Type", 0) or 0) != 3:  # msoChart
                continue
            _add(shape, source=f"Shapes[{idx}]")
        except Exception:
            continue

    found.sort(key=lambda item: (item[0], item[1]))
    return found


def _screenshot_pmi_charts(workbook_path: Path, sheet_name: str) -> dict[str, bytes]:
    """Screenshot PMI Motorized / Stationary graphs from Excel."""
    excel = None
    wb = None
    by_key: dict[str, bytes] = {}
    ordered: list[bytes] = []
    try:
        excel, wb = _open_excel_workbook(workbook_path)
        ws = _find_com_worksheet(wb, sheet_name)
        ws.Activate()
        try:
            excel.ActiveWindow.Zoom = 100
        except Exception:
            pass
        time.sleep(0.3)

        charts = _iter_sheet_charts(ws)
        print(f">>> PMI sheet: {len(charts)} Excel graph(s) to screenshot")
        scored: list[tuple[float, float, str, bytes]] = []
        for idx, (top, left, title, chart_obj) in enumerate(charts, start=1):
            label = f"PMI graph {idx} ({title or 'untitled'})"
            try:
                try:
                    chart_obj.Activate()
                except Exception:
                    try:
                        chart_obj.Select()
                    except Exception:
                        pass
                time.sleep(0.2)
                data = _screenshot_excel_chart(chart_obj, label=label)
                scored.append((top, left, title, data))
                print(f">>> Screenshotted PMI graph {idx}: {title!r} ({len(data)} bytes)")
            except Exception as exc:
                logger.warning("%s failed: %s", label, exc)
                print(f">>> WARNING: could not screenshot PMI graph {idx}: {exc}")

        for top, left, title, data in scored:
            spec = _match_pmi_chart(title)
            if spec and spec["key"] not in by_key:
                by_key[spec["key"]] = data
                print(f">>> Matched PMI screenshot to '{spec['title']}' (title {title!r})")
            else:
                ordered.append(data)

        for spec in PMI_CHART_SPECS:
            if spec["key"] in by_key:
                continue
            if not ordered:
                break
            by_key[spec["key"]] = ordered.pop(0)
            print(f">>> Assigned left→right/top→bottom PMI screenshot to '{spec['title']}'")
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


def _remove_pmi_content_shapes(slide) -> int:
    """Remove OLE table object, native charts, and prior content pictures (keep logo)."""
    removed = 0
    for shape in list(slide.shapes):
        drop = False
        st = shape.shape_type
        if st in (
            MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
            getattr(MSO_SHAPE_TYPE, "LINKED_OLE_OBJECT", MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT),
        ):
            drop = True
        elif getattr(shape, "has_chart", False):
            drop = True
        elif st == MSO_SHAPE_TYPE.PICTURE and 900_000 <= int(shape.top) < 6_800_000:
            drop = True
        if not drop:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
    return removed


def apply_pmi_workings_panels(slide, data, element: dict) -> bool:
    """Fill PMI Compliance from Workings!PMI table + graph screenshots."""
    workbook = element.get("workbook", "workings")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "contain")).lower()

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("PMI workings screenshot skipped: %s", exc)
        return False

    available = []
    try:
        available = list(data.sheet_names(workbook))
    except Exception:
        available = []

    try:
        sheet_name = resolve_sheet_name(
            workbook_bytes,
            sheet=element.get("sheet", "PMI"),
            sheet_index=element.get("sheet_index"),
            sheet_match=element.get("sheet_match", ["PMI", "PMI COMPLIANCE", "PM"]),
            sheet_match_index=int(element.get("sheet_match_index", 0) or 0),
            available=available or None,
        )
    except Exception as exc:
        logger.warning("Could not resolve Workings PMI sheet: %s", exc)
        return False

    out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    placed = 0

    # 1) Table screenshot (Regions Motorized/Stationary + Non-Motorized).
    table_png = None
    try:
        start_row, end_row, start_col, end_col = _discover_pmi_table(workbook_bytes, sheet_name)
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
            table_png, label=f"PMI table {sheet_name}", min_w=80, min_h=40, require_wide=False
        )
    except Exception as exc:
        logger.warning("PMI table capture failed: %s", exc)
        print(f">>> ERROR: PMI table screenshot failed: {exc}")
        table_png = None

    # 2) Excel graph screenshots (Motorized + Stationary).
    charts_by_key: dict[str, bytes] = {}
    if prefer_com:
        try:
            with tempfile.TemporaryDirectory(prefix="mpr_pmi_") as tmp:
                path = Path(tmp) / "workings.xlsx"
                path.write_bytes(workbook_bytes)
                charts_by_key = _screenshot_pmi_charts(path, sheet_name)
        except Exception as exc:
            logger.error("PMI Excel graph screenshots unavailable: %s", exc)
            print(f">>> ERROR: could not screenshot PMI graphs from Excel: {exc}")
    else:
        print(">>> ERROR: PMI graphs require Excel COM screenshots (prefer_excel_com=true)")

    if not table_png and not charts_by_key:
        logger.warning("No PMI table/chart screenshots from workings/%s", sheet_name)
        return False

    removed = _remove_pmi_content_shapes(slide)
    logger.info("Slide 11 cleared %s OLE/chart/picture shape(s)", removed)

    if table_png:
        left, top, width, height = PMI_TABLE_BOX
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
                debug = out_dir / f"_debug_pmi_table_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> PMI table screenshot placed from workings/{sheet_name} "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PMI table screenshot placed from workings/{sheet_name}")

    for spec in PMI_CHART_SPECS:
        png = charts_by_key.get(spec["key"])
        if png is None:
            print(f">>> WARNING: missing Excel screenshot for PMI '{spec['title']}' graph")
            continue
        left, top, width, height = PMI_CHART_BOXES[spec["key"]]
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
                debug = out_dir / f"_debug_pmi_{spec['key']}_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> PMI '{spec['title']}' graph screenshot placed "
                    f"({img.width}x{img.height}) -> {debug.name}"
                )
        except Exception:
            print(f">>> PMI '{spec['title']}' graph screenshot placed")

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> PMI Leading Issues / Action Plans cleared ({n} text box(es))")

    print(
        f"\n>>> Slide 11 PMI Compliance: placed {placed} screenshot(s) from workings/{sheet_name} "
        f"(table={'yes' if table_png else 'no'}, "
        f"graph screenshots={len(charts_by_key)}/2)\n"
    )
    return placed > 0
