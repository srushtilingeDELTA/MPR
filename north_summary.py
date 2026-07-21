"""North Scorecard Summary (PPT 14) from GSE MPR Visualizations.xlsx.

Screenshots the Scorecard Summaries tab:
  - Main scorecard block (GSE MPR + SAFETY / CX / OPS / FINANCE / PEOPLE / TOTAL
    and the listed KPI rows, including QC / Budget / OT / Hours / Lead Input)
  - Two full legend tables (not just the "Legend" header word)

Placement prefers the template's existing picture slots (largest = main,
remaining two = legends). Falls back to measured EMU boxes from the live deck
when the template pictures are already gone.
"""

from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE_TYPE

from scorecard_screenshots import (
    _fill_rgb,
    _find_sheet_name,
    _range_address,
    _remove_slide_pictures,
    _used_bounds,
    _validate_capture,
    capture_range_png,
    place_picture_on_slide,
    resolve_sheet_name,
)

logger = logging.getLogger(__name__)

# Largest template picture slot observed on North Scorecard Summary (EMU).
NORTH_MAIN_BOX = (2_208_251, 1_009_877, 7_775_497, 4_838_246)

# Fallback legend slots when the template no longer has the two smaller pictures.
NORTH_LEGEND_BOXES = (
    (314_628, 1_009_877, 1_750_000, 2_200_000),
    (314_628, 3_350_000, 1_750_000, 2_200_000),
)

CATEGORY_HEADERS = (
    "gse mpr",
    "safety",
    "customer experience",
    "operations",
    "finance",
    "people",
    "total",
)

KPI_KEYS = (
    "global injury rate",
    "ea compliance",
    "asap",
    "isr%",
    "sev",
    "pmi",
    "pmi nme",
    "qc compliance",
    "budget $000s",
    "overtime",
    "total hours",
    "lead input",
)

KPI_ALIASES: dict[str, tuple[str, ...]] = {
    "global injury rate": ("global injury rate", "gir", "injury rate"),
    "ea compliance": ("ea compliance", "eac", "ea comp"),
    "asap": ("asap",),
    "isr%": ("isr%", "isr %", "isr"),
    "sev": ("sev", "severity"),
    "pmi": ("pmi", "pm (m)", "pm motorized"),
    "pmi nme": ("pmi nme", "pm (nme)", "pm nme", "non-motorized", "non motorized"),
    "qc compliance": ("qc compliance", "qc comp", "qcc"),
    "budget $000s": (
        "budget $000s",
        "budget $000",
        "budget ($000s)",
        "budget ($000)",
        "budget",
    ),
    "overtime": ("overtime", "ot", "o.t."),
    "total hours": ("total hours", "hours", "total hrs", "tot hours"),
    "lead input": ("lead input", "lead inputs", "leadership input"),
}

REQUIRED_BOTTOM_KPIS = (
    "qc compliance",
    "budget $000s",
    "overtime",
    "total hours",
    "lead input",
)

LEGEND_TOKENS = (
    "legend",
    "color key",
    "status key",
    "score key",
    "key:",
)

# Minimum footprint so we never paste only the word "Legend".
LEGEND_MIN_EXTRA_ROWS = 3
LEGEND_MIN_EXTRA_COLS = 2
LEGEND_MAX_ROWS = 12
LEGEND_MAX_COLS = 6


@dataclass
class RangeBlock:
    """A rectangular Excel capture region (1-based inclusive)."""

    label: str
    start_row: int
    end_row: int
    start_col: int
    end_col: int

    @property
    def range_address(self) -> str:
        return _range_address(self.start_row, self.end_row, self.start_col, self.end_col)

    @property
    def row_count(self) -> int:
        return self.end_row - self.start_row + 1

    @property
    def col_count(self) -> int:
        return self.end_col - self.start_col + 1


@dataclass
class NorthSummaryLayout:
    main: RangeBlock
    legends: list[RangeBlock]


def _cell_str(ws, row: int, col: int) -> str:
    val = ws.cell(row, col).value
    if val is None:
        return ""
    return str(val).strip()


def _norm(text: str) -> str:
    cleaned = (
        str(text or "")
        .replace("\xa0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )
    cleaned = re.sub(r"[()]", " ", cleaned)
    return re.sub(r"\s+", " ", cleaned).casefold().strip()


def _cell_has_fill(ws, row: int, col: int) -> bool:
    try:
        rgb = _fill_rgb(ws.cell(row, col))
    except Exception:
        return False
    if rgb is None:
        return False
    return not (rgb[0] >= 245 and rgb[1] >= 245 and rgb[2] >= 245)


def _cell_interesting(ws, row: int, col: int) -> bool:
    """True when a cell has text or a non-white fill (legend color swatches)."""
    return bool(_cell_str(ws, row, col)) or _cell_has_fill(ws, row, col)


def _match_category(text: str, token: str) -> bool:
    n = _norm(text)
    if not n:
        return False
    if token == "total":
        return n == "total" or (n.startswith("total ") and "hour" not in n)
    if token == "safety":
        return n == "safety" or n.startswith("safety")
    return token == n or token in n or n in token


def _match_kpi(text: str, key: str) -> bool:
    n = _norm(text)
    if not n:
        return False
    for alias in KPI_ALIASES.get(key, (key,)):
        a = _norm(alias)
        if not a:
            continue
        # Short aliases (ot, sev, qc) must be whole-token matches.
        if len(a) <= 3:
            tokens = set(re.findall(r"[a-z0-9%]+", n))
            if a not in tokens and n != a:
                continue
        else:
            hit = n == a or a in n or a.replace(" ", "") in n.replace(" ", "")
            if not hit:
                continue
        if key == "pmi" and "nme" in n:
            continue
        if key == "lead input" and "leading" in n:
            continue
        if key == "isr%" and n in {"severity", "sev"}:
            continue
        if key == "overtime" and "total" in n and "overtime" not in n and "ot" not in n.split():
            continue
        return True
    return False


def _find_label_cells(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> dict[str, tuple[int, int]]:
    """Map canonical label key -> first (row, col) hit."""
    found: dict[str, tuple[int, int]] = {}
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            text = _cell_str(ws, row, col)
            if not text:
                continue
            for token in CATEGORY_HEADERS:
                if token not in found and _match_category(text, token):
                    found[token] = (row, col)
            for key in KPI_KEYS:
                if key not in found and _match_kpi(text, key):
                    found[key] = (row, col)
    return found


def _picture_boxes(slide) -> list[tuple[int, int, int, int]]:
    boxes: list[tuple[int, int, int, int]] = []
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        boxes.append((int(shape.left), int(shape.top), int(shape.width), int(shape.height)))
    boxes.sort(key=lambda b: b[2] * b[3], reverse=True)
    return boxes


def _resolve_placement_boxes(
    slide,
    element: dict,
) -> tuple[tuple[int, int, int, int], list[tuple[int, int, int, int]]]:
    if element.get("main_box"):
        main = tuple(int(x) for x in element["main_box"])  # type: ignore[assignment]
    else:
        main = None

    legend_boxes: list[tuple[int, int, int, int]] | None = None
    raw_legends = element.get("legend_boxes")
    if raw_legends:
        legend_boxes = [tuple(int(x) for x in box) for box in raw_legends]

    slots = _picture_boxes(slide)
    if main is None and slots:
        main = slots[0]
    if legend_boxes is None and len(slots) >= 3:
        legend_boxes = sorted(slots[1:3], key=lambda b: (b[1], b[0]))
    elif legend_boxes is None and len(slots) == 2:
        legend_boxes = [slots[1], NORTH_LEGEND_BOXES[1]]

    if main is None:
        if all(element.get(k) is not None for k in ("left", "top", "max_width", "max_height")):
            main = (
                int(element["left"]),
                int(element["top"]),
                int(element["max_width"]),
                int(element["max_height"]),
            )
        else:
            main = NORTH_MAIN_BOX

    if not legend_boxes:
        legend_boxes = list(NORTH_LEGEND_BOXES)
    return main, legend_boxes[:2]


def _expand_main_block(
    ws,
    seed_rows: list[int],
    seed_cols: list[int],
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    legend_anchors: list[tuple[int, int]],
    required_end_row: int | None,
) -> tuple[int, int, int, int]:
    """Expand to the full scorecard; never truncate below required KPI rows."""
    if not seed_rows or not seed_cols:
        return min_row, max_row, min_col, max_col

    start_row = max(min_row, min(seed_rows) - 1)
    end_row = min(max_row, max(seed_rows))
    start_col = max(min_col, min(seed_cols))
    end_col = min(max_col, max(seed_cols) + 1)

    if required_end_row is not None:
        end_row = max(end_row, min(max_row, required_end_row))

    # Side legends (right of the scorecard) must not be pulled into main.
    kpi_right = max(seed_cols)
    side_legend_cols = sorted(c for r, c in legend_anchors if c > kpi_right + 1)
    col_cap = min(side_legend_cols) - 1 if side_legend_cols else max_col

    # Only legends BELOW the required KPI block may cap the main range.
    if required_end_row is not None:
        below = [r for r, c in legend_anchors if r > required_end_row and c <= col_cap]
        if below and end_row >= min(below):
            end_row = max(required_end_row, min(below) - 1)

    empty_streak = 0
    cursor = end_row
    while cursor < max_row and empty_streak < 3:
        nxt = cursor + 1
        if any(
            r == nxt and c <= col_cap and (required_end_row is None or nxt > required_end_row)
            for r, c in legend_anchors
        ):
            break
        if any(_cell_interesting(ws, nxt, c) for c in range(start_col, min(end_col, col_cap) + 1)):
            end_row = nxt
            cursor = nxt
            empty_streak = 0
        else:
            empty_streak += 1
            cursor = nxt

    if required_end_row is not None:
        end_row = max(end_row, min(max_row, required_end_row))

    while end_col < max_col and end_col < col_cap:
        nxt = end_col + 1
        if nxt > col_cap:
            break
        if any(_cell_interesting(ws, r, nxt) for r in range(start_row, end_row + 1)):
            end_col = nxt
            continue
        if nxt + 1 <= col_cap and any(
            _cell_interesting(ws, r, nxt + 1) for r in range(start_row, end_row + 1)
        ):
            end_col = nxt + 1
            continue
        break

    end_col = min(end_col, col_cap)

    try:
        for merged in ws.merged_cells.ranges:
            if (
                merged.max_row >= start_row
                and merged.min_row <= end_row
                and merged.max_col >= start_col
                and merged.min_col <= end_col
            ):
                start_row = min(start_row, int(merged.min_row))
                end_row = max(end_row, int(merged.max_row))
                start_col = min(start_col, int(merged.min_col))
                end_col = min(col_cap, max(end_col, int(merged.max_col)))
    except Exception:
        pass

    if required_end_row is not None:
        end_row = max(end_row, min(max_row, required_end_row))

    return start_row, max(start_row, end_row), start_col, max(start_col, end_col)


def _is_legend_anchor(text: str) -> bool:
    n = _norm(text)
    if not n:
        return False
    return any(tok in n for tok in LEGEND_TOKENS)


def _expand_legend_table(
    ws,
    anchor_row: int,
    anchor_col: int,
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    exclude: tuple[int, int, int, int] | None = None,
    other_legend_anchors: list[tuple[int, int]] | None = None,
) -> RangeBlock | None:
    """Grow a legend from its header to the full small table (text + color fills)."""

    def in_exclude(r: int, c: int) -> bool:
        if exclude is None:
            return False
        er0, er1, ec0, ec1 = exclude
        return er0 <= r <= er1 and ec0 <= c <= ec1

    others = [
        (r, c)
        for r, c in (other_legend_anchors or [])
        if not (r == anchor_row and c == anchor_col)
    ]

    def hits_other_legend(r: int, c: int) -> bool:
        return any(r == or_ and c == oc for or_, oc in others)

    start_row, end_row = anchor_row, anchor_row
    start_col, end_col = anchor_col, anchor_col

    # Always include a minimum body under/beside the Legend header.
    end_row = min(max_row, anchor_row + LEGEND_MIN_EXTRA_ROWS)
    end_col = min(max_col, anchor_col + LEGEND_MIN_EXTRA_COLS)

    # Do not let the minimum footprint cross into a neighboring Legend header.
    for r, c in others:
        if r == anchor_row and c > anchor_col:
            end_col = min(end_col, c - 1)
        if c == anchor_col and r > anchor_row:
            end_row = min(end_row, r - 1)

    for _ in range(60):
        grew = False
        if end_row < max_row and (end_row - start_row + 1) < LEGEND_MAX_ROWS:
            r = end_row + 1
            if any(hits_other_legend(r, c) for c in range(start_col, end_col + 1)):
                pass
            elif any(
                _cell_interesting(ws, r, c) and not in_exclude(r, c)
                for c in range(start_col, end_col + 1)
            ):
                end_row = r
                grew = True
        if end_col < max_col and (end_col - start_col + 1) < LEGEND_MAX_COLS:
            c = end_col + 1
            if any(hits_other_legend(r, c) for r in range(start_row, end_row + 1)):
                pass
            elif any(
                _cell_interesting(ws, r, c) and not in_exclude(r, c)
                for r in range(start_row, end_row + 1)
            ):
                end_col = c
                grew = True
        if start_col > min_col and (end_col - start_col + 1) < LEGEND_MAX_COLS:
            c = start_col - 1
            if any(hits_other_legend(r, c) for r in range(start_row, end_row + 1)):
                pass
            elif any(
                _cell_interesting(ws, r, c) and not in_exclude(r, c)
                for r in range(start_row, end_row + 1)
            ):
                start_col = c
                grew = True
        if start_row > min_row and (end_row - start_row + 1) < LEGEND_MAX_ROWS:
            r = start_row - 1
            if any(hits_other_legend(r, c) for c in range(start_col, end_col + 1)):
                pass
            elif any(
                _cell_interesting(ws, r, c) and not in_exclude(r, c)
                for c in range(start_col, end_col + 1)
            ):
                if not (exclude is not None and in_exclude(r, start_col)):
                    start_row = r
                    grew = True
        if not grew:
            break

    min_end_row = min(max_row, anchor_row + LEGEND_MIN_EXTRA_ROWS)
    min_end_col = min(max_col, anchor_col + LEGEND_MIN_EXTRA_COLS)
    for r, c in others:
        if r == anchor_row and c > anchor_col:
            min_end_col = min(min_end_col, c - 1)
        if c == anchor_col and r > anchor_row:
            min_end_row = min(min_end_row, r - 1)

    while end_row > max(anchor_row, min_end_row) and not any(
        _cell_interesting(ws, end_row, c) for c in range(start_col, end_col + 1)
    ):
        end_row -= 1
    while end_col > max(anchor_col, min_end_col) and not any(
        _cell_interesting(ws, r, end_col) for r in range(start_row, end_row + 1)
    ):
        end_col -= 1

    if end_row < start_row or end_col < start_col:
        return None
    if (end_row - start_row + 1) * (end_col - start_col + 1) > 500:
        return None

    return RangeBlock(
        label=f"legend@{get_column_letter(start_col)}{start_row}",
        start_row=start_row,
        end_row=end_row,
        start_col=start_col,
        end_col=end_col,
    )


def _orphan_tables_outside_main(
    ws,
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    main: tuple[int, int, int, int],
    already: list[RangeBlock],
) -> list[RangeBlock]:
    """Find small filled regions outside the main scorecard (likely legends)."""
    mr0, mr1, mc0, mc1 = main
    already_keys = {(b.start_row, b.end_row, b.start_col, b.end_col) for b in already}
    candidates: list[RangeBlock] = []

    search_cells: list[tuple[int, int]] = []
    for row in range(mr1 + 1, max_row + 1):
        for col in range(min_col, max_col + 1):
            if _cell_interesting(ws, row, col):
                search_cells.append((row, col))
                break
    for col in list(range(min_col, mc0)) + list(range(mc1 + 1, max_col + 1)):
        for row in range(min_row, max_row + 1):
            if _cell_interesting(ws, row, col) and not (mr0 <= row <= mr1 and mc0 <= col <= mc1):
                search_cells.append((row, col))
                break

    for row, col in search_cells:
        block = _expand_legend_table(
            ws,
            row,
            col,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            exclude=main,
        )
        if block is None:
            continue
        key = (block.start_row, block.end_row, block.start_col, block.end_col)
        if key in already_keys:
            continue
        if block.row_count < 2 and block.col_count < 2:
            continue
        overlap_rows = max(0, min(block.end_row, mr1) - max(block.start_row, mr0) + 1)
        if overlap_rows > (block.end_row - block.start_row + 1) * 0.5:
            continue
        already_keys.add(key)
        candidates.append(block)

    candidates.sort(key=lambda b: (b.start_row, b.start_col))
    return candidates[: max(0, 2 - len(already))]


def detect_north_summary_layout(workbook_bytes: bytes, sheet_name: str) -> NorthSummaryLayout:
    """Detect main scorecard + two full legend ranges on Scorecard Summaries."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    match = _find_sheet_name(list(wb.sheetnames), sheet_name)
    if match is None:
        wb.close()
        raise ValueError(f"Sheet {sheet_name!r} not found. Available: {wb.sheetnames}")
    ws = wb[match]
    min_row, max_row, min_col, max_col = _used_bounds(ws)

    legend_anchors: list[tuple[int, int]] = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if _is_legend_anchor(_cell_str(ws, row, col)):
                legend_anchors.append((row, col))

    labels = _find_label_cells(ws, min_row, max_row, min_col, max_col)
    category_hits = [labels[t] for t in CATEGORY_HEADERS if t in labels]
    kpi_hits = [labels[k] for k in KPI_KEYS if k in labels]
    bottom_hits = [labels[k] for k in REQUIRED_BOTTOM_KPIS if k in labels]
    required_end_row = max((r for r, _ in bottom_hits), default=None)
    if required_end_row is None and kpi_hits:
        required_end_row = max(r for r, _ in kpi_hits)

    missing_bottom = [k for k in REQUIRED_BOTTOM_KPIS if k not in labels]
    if missing_bottom:
        logger.warning(
            "North summary missing KPI labels on %s: %s (found=%s)",
            match,
            missing_bottom,
            sorted(k for k in labels if k in KPI_KEYS),
        )

    if category_hits or kpi_hits:
        seed_rows = [r for r, _ in category_hits + kpi_hits]
        seed_cols = [c for _, c in category_hits + kpi_hits]
        start_row, end_row, start_col, end_col = _expand_main_block(
            ws,
            seed_rows,
            seed_cols,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            legend_anchors=legend_anchors,
            required_end_row=required_end_row,
        )
    else:
        logger.warning(
            "North summary: category/KPI labels not found on %s — using used range",
            match,
        )
        start_row, end_row, start_col, end_col = min_row, max_row, min_col, max_col

    main = RangeBlock(
        label="north_scorecard",
        start_row=start_row,
        end_row=max(start_row, end_row),
        start_col=start_col,
        end_col=end_col,
    )

    legends: list[RangeBlock] = []
    seen: set[tuple[int, int, int, int]] = set()
    for row, col in sorted(legend_anchors, key=lambda rc: (rc[0], rc[1])):
        block = _expand_legend_table(
            ws,
            row,
            col,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            exclude=(main.start_row, main.end_row, main.start_col, main.end_col)
            if row > main.end_row or col > main.end_col
            else None,
            other_legend_anchors=legend_anchors,
        )
        if block is None:
            continue
        if block.row_count < 2 and block.col_count < 2:
            continue
        key = (block.start_row, block.end_row, block.start_col, block.end_col)
        if key in seen:
            continue
        if any(abs(block.start_row - s[0]) <= 1 and abs(block.start_col - s[2]) <= 1 for s in seen):
            continue
        seen.add(key)
        legends.append(block)
        if len(legends) >= 2:
            break

    if len(legends) < 2:
        extras = _orphan_tables_outside_main(
            ws,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            main=(main.start_row, main.end_row, main.start_col, main.end_col),
            already=legends,
        )
        for block in extras:
            if len(legends) >= 2:
                break
            legends.append(block)

    wb.close()
    logger.info(
        "North summary layout on %s: main=%s legends=%s labels=%s missing_bottom=%s",
        match,
        main.range_address,
        [b.range_address for b in legends],
        sorted(labels.keys()),
        missing_bottom,
    )
    print(
        f">>> North layout {match}: main={main.range_address} "
        f"legends={[b.range_address for b in legends]} "
        f"bottom_kpis={[k for k in REQUIRED_BOTTOM_KPIS if k in labels]}"
    )
    return NorthSummaryLayout(main=main, legends=legends)


def _capture_block(
    workbook_bytes: bytes,
    sheet_name: str,
    block: RangeBlock,
    *,
    prefer_com: bool,
) -> bytes | None:
    try:
        png = capture_range_png(
            workbook_bytes,
            sheet_name,
            block.start_row,
            block.end_row,
            block.start_col,
            block.end_col,
            prefer_excel_com=prefer_com,
        )
        return _validate_capture(
            png,
            label=f"north {block.label} {sheet_name}!{block.range_address}",
            min_w=30,
            min_h=15,
            require_wide=False,
        )
    except Exception as exc:
        msg = str(exc).casefold()
        if "mostly blank" in msg or "too small" in msg:
            try:
                from scorecard_screenshots import _capture_via_render, _png_from_pil

                raw = _capture_via_render(
                    workbook_bytes,
                    sheet_name,
                    block.start_row,
                    block.end_row,
                    block.start_col,
                    block.end_col,
                )
                with Image.open(io.BytesIO(raw)) as img:
                    if img.width >= 30 and img.height >= 15:
                        logger.warning(
                            "North capture %s accepted after lenient fallback (%sx%s): %s",
                            block.range_address,
                            img.width,
                            img.height,
                            exc,
                        )
                        return _png_from_pil(img)
            except Exception:
                pass
        logger.warning("North capture failed for %s: %s", block.range_address, exc)
        return None


def apply_north_summary_panels(slide, data, element: dict) -> bool:
    """Screenshot Visualizations Scorecard Summaries onto North Scorecard Summary."""
    workbook = element.get("workbook", "visualizations")
    prefer_com = bool(element.get("prefer_excel_com", True))
    fit = str(element.get("fit", "fill")).lower()
    replace_existing = bool(element.get("replace_existing_pictures", True))

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("North summary skipped: %s", exc)
        print(f">>> North summary skipped — visualizations workbook missing ({exc})")
        return False

    available = []
    try:
        available = list(data.sheet_names(workbook))
    except Exception:
        available = []

    try:
        sheet_name = resolve_sheet_name(
            workbook_bytes,
            sheet=element.get("sheet"),
            sheet_index=element.get("sheet_index"),
            sheet_match=element.get("sheet_match", ["scorecard summar"]),
            sheet_match_index=int(element.get("sheet_match_index", 0) or 0),
            available=available or None,
        )
    except Exception as exc:
        logger.warning("Could not resolve Scorecard Summaries sheet: %s", exc)
        return False

    try:
        layout = detect_north_summary_layout(workbook_bytes, sheet_name)
    except Exception as exc:
        logger.warning("North summary layout detection failed: %s", exc)
        return False

    main_png = _capture_block(workbook_bytes, sheet_name, layout.main, prefer_com=prefer_com)
    legend_pngs: list[bytes] = []
    for legend in layout.legends:
        png = _capture_block(workbook_bytes, sheet_name, legend, prefer_com=prefer_com)
        if png:
            legend_pngs.append(png)

    if not main_png:
        print(f">>> FAILED north summary main capture from {workbook}/{sheet_name}")
        return False

    main_box, legend_boxes = _resolve_placement_boxes(slide, element)

    if replace_existing:
        removed = _remove_slide_pictures(slide)
        if removed:
            logger.info("Removed %s template picture(s) before North summary placement", removed)
            print(f">>> Removed {removed} template picture(s) for North summary")

    out_dir = Path(getattr(data.store, "base_dir", Path("."))) / "output"
    out_dir.mkdir(parents=True, exist_ok=True)

    place_picture_on_slide(
        slide,
        main_png,
        left=main_box[0],
        top=main_box[1],
        max_width=main_box[2],
        max_height=main_box[3],
        fit=fit,
    )
    placed = 1
    try:
        with Image.open(io.BytesIO(main_png)) as img:
            debug = out_dir / f"_debug_north_main_{img.width}x{img.height}.png"
            img.save(debug)
            print(
                f">>> North main scorecard from {workbook}/{sheet_name}!"
                f"{layout.main.range_address} ({img.width}x{img.height}) "
                f"box={main_box} -> {debug.name}"
            )
    except Exception:
        print(f">>> North main scorecard placed from {workbook}/{sheet_name}")

    for idx, png in enumerate(legend_pngs[:2]):
        box = legend_boxes[idx] if idx < len(legend_boxes) else NORTH_LEGEND_BOXES[min(idx, 1)]
        place_picture_on_slide(
            slide,
            png,
            left=box[0],
            top=box[1],
            max_width=box[2],
            max_height=box[3],
            fit=fit,
        )
        placed += 1
        legend = layout.legends[idx]
        try:
            with Image.open(io.BytesIO(png)) as img:
                debug = out_dir / f"_debug_north_legend{idx + 1}_{img.width}x{img.height}.png"
                img.save(debug)
                print(
                    f">>> North legend {idx + 1} from {sheet_name}!{legend.range_address} "
                    f"({img.width}x{img.height}) box={box} -> {debug.name}"
                )
        except Exception:
            print(f">>> North legend {idx + 1} placed")

    if len(legend_pngs) < 2:
        print(
            f">>> WARNING: expected 2 legend screenshots, captured {len(legend_pngs)} "
            f"from {sheet_name}"
        )

    print(
        f"\n>>> Slide 14 North Scorecard Summary: placed {placed} screenshot(s) "
        f"from {workbook}/{sheet_name}\n"
    )
    return placed > 0
