"""EA Compliance / ASAP slide panels from New GSE MPR Workings.xlsx.

Captures the EAC and ASAP tabs as screenshots and places them into the
left content slot on PPT slide 6 (replacing the embedded OLE object).
Leading Issues / Action Plan stay as empty editable text boxes on the right.
"""

from __future__ import annotations

import logging
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import io

from gir_panels import clear_leading_action_narrative
from scorecard_screenshots import (
    capture_sheet_png,
    place_picture_on_slide,
    resolve_sheet_name,
    _validate_capture,
)

logger = logging.getLogger(__name__)

# Left content slot occupied by the template OLE object on slide 6.
EA_CONTENT_LEFT = 314628
EA_CONTENT_TOP = 916047
EA_CONTENT_WIDTH = 7924800
EA_CONTENT_HEIGHT = 3702050
EA_PANEL_GAP = 90_000

DEFAULT_PANELS = [
    {
        "key": "eac",
        "sheet": "EAC",
        "sheet_match": ["EAC"],
        "label": "EAC",
    },
    {
        "key": "asap",
        "sheet": "ASAP",
        "sheet_match": ["ASAP"],
        "label": "ASAP",
    },
]


def _remove_content_ole_and_pictures(slide, *, band_top: int, band_bottom: int) -> int:
    """Remove OLE objects and pictures in the left content band (keep bottom logo)."""
    removed = 0
    for shape in list(slide.shapes):
        st = shape.shape_type
        is_ole = st in (
            MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT,
            getattr(MSO_SHAPE_TYPE, "LINKED_OLE_OBJECT", MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT),
        )
        is_pic = st == MSO_SHAPE_TYPE.PICTURE
        if not is_ole and not is_pic:
            continue
        top = int(shape.top)
        bottom = top + int(shape.height)
        # Keep logo / footer images below the content band.
        if top >= band_bottom - 50_000:
            continue
        if bottom < band_top:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
        logger.info("Removed slide 6 content shape %s (%s)", shape.name, st)
    return removed


def _panel_boxes(count: int) -> list[tuple[int, int, int, int]]:
    """Stack N panel boxes inside the EA content slot."""
    if count <= 0:
        return []
    if count == 1:
        return [(EA_CONTENT_LEFT, EA_CONTENT_TOP, EA_CONTENT_WIDTH, EA_CONTENT_HEIGHT)]

    usable = EA_CONTENT_HEIGHT - EA_PANEL_GAP * (count - 1)
    height = max(200_000, usable // count)
    boxes = []
    top = EA_CONTENT_TOP
    for idx in range(count):
        boxes.append((EA_CONTENT_LEFT, top, EA_CONTENT_WIDTH, height))
        top += height + EA_PANEL_GAP
    # Absorb rounding into the last box so we still fill the slot.
    if boxes:
        last_left, last_top, last_w, _ = boxes[-1]
        boxes[-1] = (
            last_left,
            last_top,
            last_w,
            max(height, EA_CONTENT_TOP + EA_CONTENT_HEIGHT - last_top),
        )
    return boxes


def apply_ea_asap_workings_panels(slide, data, element: dict) -> bool:
    """Screenshot Workings EAC + ASAP tabs into the slide 6 content slot."""
    workbook = element.get("workbook", "workings")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "contain")).lower()
    panels = element.get("panels") or DEFAULT_PANELS

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("EA/ASAP workings screenshot skipped: %s", exc)
        return False

    available = []
    try:
        available = list(data.sheet_names(workbook))
    except Exception:
        available = []

    captures: list[tuple[str, bytes]] = []
    for panel in panels:
        label = str(panel.get("label") or panel.get("key") or panel.get("sheet") or "panel")
        try:
            sheet_name = resolve_sheet_name(
                workbook_bytes,
                sheet=panel.get("sheet"),
                sheet_index=panel.get("sheet_index"),
                sheet_match=panel.get("sheet_match"),
                sheet_match_index=int(panel.get("sheet_match_index", 0) or 0),
                available=available or None,
            )
        except Exception as exc:
            logger.warning("Could not resolve Workings sheet for %s: %s", label, exc)
            continue

        try:
            png = capture_sheet_png(
                workbook_bytes,
                sheet_name,
                prefer_excel_com=prefer_com,
                max_rows=panel.get("max_rows", element.get("max_rows", 80)),
                max_cols=panel.get("max_cols", element.get("max_cols", 25)),
            )
            png = _validate_capture(
                png,
                label=f"Workings {sheet_name}",
                min_w=80,
                min_h=40,
                require_wide=False,
            )
        except Exception as exc:
            logger.warning("Capture failed for Workings!%s (%s): %s", sheet_name, label, exc)
            continue

        captures.append((f"{label}:{sheet_name}", png))
        try:
            with Image.open(io.BytesIO(png)) as img:
                out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
                out_dir.mkdir(parents=True, exist_ok=True)
                debug = out_dir / f"_debug_ea_{label}_{img.width}x{img.height}.png"
                img.save(debug)
                print(f">>> EA/ASAP panel '{label}' captured from workings/{sheet_name} ({img.width}x{img.height})")
        except Exception:
            print(f">>> EA/ASAP panel '{label}' captured from workings/{sheet_name}")

    if not captures:
        logger.warning("No EAC/ASAP panels captured from workings")
        return False

    band_top = EA_CONTENT_TOP - 50_000
    band_bottom = EA_CONTENT_TOP + EA_CONTENT_HEIGHT + 50_000
    removed = _remove_content_ole_and_pictures(slide, band_top=band_top, band_bottom=band_bottom)
    logger.info("Slide 6 cleared %s OLE/picture shape(s) in content band", removed)

    boxes = _panel_boxes(len(captures))
    placed = 0
    for (label, png), box in zip(captures, boxes):
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
        logger.info("Placed EA/ASAP panel %s at (%s,%s,%s,%s)", label, left, top, width, height)

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> EA/ASAP Leading Issues / Action Plan cleared ({n} text box(es))")

    print(f"\n>>> Slide 6 EA/ASAP: placed {placed} screenshot(s) from workings (EAC/ASAP)\n")
    return placed > 0
