"""EA Compliance / ASAP slide from New GSE MPR Workings.xlsx.

Captures the EAC and ASAP tabs (once each), stitches them into a single
image, and places that image once into the left content slot on slide 6.
Leading Issues / Action Plan stay as empty editable text boxes on the right.
"""

from __future__ import annotations

import hashlib
import io
import logging
from pathlib import Path

from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from gir_panels import clear_leading_action_narrative
from scorecard_screenshots import (
    _stitch_vertical_many,
    _validate_capture,
    capture_sheet_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Left content slot on slide 6 — sized to stay clear of:
# title (~top), Leading Issues column (~left 8.42M), and bottom logo (~top 6.35M).
EA_CONTENT_LEFT = 300_000
EA_CONTENT_TOP = 880_000
EA_CONTENT_WIDTH = 7_900_000
EA_CONTENT_HEIGHT = 5_200_000

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


def _remove_content_ole_and_pictures(slide, *, band_top: int, band_bottom: int, band_right: int) -> int:
    """Remove OLE/pictures in the left content band (keep bottom logo + right narrative)."""
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
        left = int(shape.left)
        top = int(shape.top)
        bottom = top + int(shape.height)
        # Keep logo / footer images below the content band.
        if top >= band_bottom - 50_000:
            continue
        # Keep anything primarily in the Leading Issues column.
        if left >= band_right - 100_000:
            continue
        if bottom < band_top:
            continue
        shape._element.getparent().remove(shape._element)
        removed += 1
        logger.info("Removed slide 6 content shape %s (%s)", shape.name, st)
    return removed


def _png_fingerprint(png: bytes) -> str:
    return hashlib.sha1(png).hexdigest()


def _capture_unique_panels(
    workbook_bytes: bytes,
    available: list[str],
    panels: list[dict],
    *,
    prefer_com: bool,
    max_rows: int,
    max_cols: int,
    out_dir: Path,
) -> list[tuple[str, bytes]]:
    """Capture each requested panel once; skip duplicate sheet names / identical images."""
    captures: list[tuple[str, bytes]] = []
    used_sheets: set[str] = set()
    used_fingerprints: set[str] = set()

    for panel in panels:
        label = str(panel.get("label") or panel.get("key") or panel.get("sheet") or "panel")
        remaining = [n for n in available if n not in used_sheets]
        try:
            sheet_name = resolve_sheet_name(
                workbook_bytes,
                sheet=panel.get("sheet"),
                sheet_index=panel.get("sheet_index"),
                sheet_match=panel.get("sheet_match"),
                sheet_match_index=int(panel.get("sheet_match_index", 0) or 0),
                available=remaining or available or None,
            )
        except Exception as exc:
            logger.warning("Could not resolve Workings sheet for %s: %s", label, exc)
            continue

        if sheet_name in used_sheets:
            logger.info("Skipping %s — sheet %s already captured", label, sheet_name)
            print(f">>> EA/ASAP skip '{label}': sheet {sheet_name!r} already used")
            continue

        try:
            png = capture_sheet_png(
                workbook_bytes,
                sheet_name,
                prefer_excel_com=prefer_com,
                max_rows=panel.get("max_rows", max_rows),
                max_cols=panel.get("max_cols", max_cols),
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

        fp = _png_fingerprint(png)
        if fp in used_fingerprints:
            logger.info("Skipping %s — identical image already captured", label)
            print(f">>> EA/ASAP skip '{label}': duplicate image of a prior panel")
            used_sheets.add(sheet_name)
            continue

        used_sheets.add(sheet_name)
        used_fingerprints.add(fp)
        captures.append((f"{label}:{sheet_name}", png))

        try:
            with Image.open(io.BytesIO(png)) as img:
                out_dir.mkdir(parents=True, exist_ok=True)
                debug = out_dir / f"_debug_ea_{label}_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> EA/ASAP panel '{label}' captured from workings/{sheet_name} "
                    f"({img.width}x{img.height})"
                )
        except Exception:
            print(f">>> EA/ASAP panel '{label}' captured from workings/{sheet_name}")

    return captures


def apply_ea_asap_workings_panels(slide, data, element: dict) -> bool:
    """Screenshot Workings EAC + ASAP once and place a single large image on slide 6."""
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

    out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
    captures = _capture_unique_panels(
        workbook_bytes,
        available,
        panels,
        prefer_com=prefer_com,
        max_rows=int(element.get("max_rows", 80) or 80),
        max_cols=int(element.get("max_cols", 25) or 25),
        out_dir=out_dir,
    )
    if not captures:
        logger.warning("No EAC/ASAP panels captured from workings")
        return False

    # One image only — stitch distinct panels, never place the same table twice.
    if len(captures) == 1:
        final_png = captures[0][1]
        source_label = captures[0][0]
    else:
        final_png = _stitch_vertical_many([png for _, png in captures])
        source_label = " + ".join(label for label, _ in captures)
        try:
            with Image.open(io.BytesIO(final_png)) as img:
                debug = out_dir / f"_debug_ea_stitched_{img.width}x{img.height}.png"
                img.save(debug)
                print(f">>> EA/ASAP stitched panel ({img.width}x{img.height}) from {source_label}")
        except Exception:
            pass

    band_top = EA_CONTENT_TOP - 50_000
    band_bottom = EA_CONTENT_TOP + EA_CONTENT_HEIGHT + 50_000
    band_right = EA_CONTENT_LEFT + EA_CONTENT_WIDTH
    removed = _remove_content_ole_and_pictures(
        slide,
        band_top=band_top,
        band_bottom=band_bottom,
        band_right=band_right,
    )
    logger.info("Slide 6 cleared %s OLE/picture shape(s) in content band", removed)

    place_picture_on_slide(
        slide,
        final_png,
        left=EA_CONTENT_LEFT,
        top=EA_CONTENT_TOP,
        max_width=EA_CONTENT_WIDTH,
        max_height=EA_CONTENT_HEIGHT,
        fit=fit,
    )
    print(
        f"\n>>> Slide 6 EA/ASAP: placed 1 screenshot "
        f"({source_label}) in left content slot "
        f"{EA_CONTENT_WIDTH}x{EA_CONTENT_HEIGHT} EMU, fit={fit}\n"
    )

    if bool(element.get("clear_narrative", True)):
        n = clear_leading_action_narrative(slide)
        print(f">>> EA/ASAP Leading Issues / Action Plan cleared ({n} text box(es))")

    return True
