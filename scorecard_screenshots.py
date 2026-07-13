"""Capture System scorecard sections from Excel and prepare slide images.

Primary path (Windows + Excel): CopyPicture via COM for true screenshots.
Fallback (any OS): render openpyxl cell fills/text to PNG with Pillow.
"""

from __future__ import annotations

import io
import logging
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image, ImageDraw, ImageFont

logger = logging.getLogger(__name__)

# Content area under the System Scorecard title on slides 3/4 (EMU).
DEFAULT_LEFT = 167148
DEFAULT_TOP = 720000
DEFAULT_WIDTH = 11800000
DEFAULT_HEIGHT = 5900000


@dataclass
class Section:
    """One category block on the System scorecard (e.g. Safety & Security)."""

    index: int
    title: str
    start_row: int  # 1-based inclusive (category body start)
    end_row: int  # 1-based inclusive
    start_col: int  # 1-based inclusive
    end_col: int  # 1-based inclusive
    is_black: bool = False

    @property
    def range_address(self) -> str:
        return (
            f"{get_column_letter(self.start_col)}{self.start_row}:"
            f"{get_column_letter(self.end_col)}{self.end_row}"
        )


@dataclass
class SystemLayout:
    """Detected System sheet layout: shared month header + category sections."""

    header_start_row: int
    header_end_row: int
    start_col: int
    end_col: int
    sections: list[Section]

    def capture_bounds_for(self, sections: list[Section]) -> tuple[int, int, int, int]:
        """Body bounds for selected sections (header handled separately when needed)."""
        if not sections:
            return self.header_start_row, self.header_end_row, self.start_col, self.end_col
        return (
            min(s.start_row for s in sections),
            max(s.end_row for s in sections),
            self.start_col,
            self.end_col,
        )

    def needs_header_stitch(self, sections: list[Section]) -> bool:
        """True when selected body does not sit directly under the month header."""
        if not sections:
            return False
        body_start = min(s.start_row for s in sections)
        return body_start > self.header_end_row + 1


MONTH_HEADER_TOKENS = {
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "SEPT",
    "OCT",
    "NOV",
    "DEC",
    "YTD",
    "YE",
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
}

# Ordered System tab categories used when left-rail labels are shapes/images
# (openpyxl cannot always read those cells, but TOTAL SCORE rows still mark boundaries).
DEFAULT_SYSTEM_CATEGORIES = [
    "Safety & Security",
    "Customer Experience",
    "Operations",
    "People",
    "Finance",
    "Overall and Opportunities",
]

# KPI names that must NEVER be treated as category section titles.
KNOWN_KPI_NAMES = {
    "budget",
    "total hours",
    "hours",
    "ot",
    "psych. safety",
    "psych safety",
    "psychological safety",
    "global injury rate",
    "gir",
    "ea compliance",
    "asap",
    "in-service rate",
    "vos (stationary)",
    "vos",
    "pmi",
    "critical jam",
    "clear time",
    "mbr",
    "engagement",
    "retention",
    "training",
}

METRIC_LABELS = {
    "plan",
    "actual",
    "percent",
    "points",
    "score",
    "weight",
    "kpi",
    "total score",
}

CATEGORY_KEYWORDS = (
    "safety & security",
    "customer experience",
    "operations",
    "people",
    "finance",
    "financial",
    "overall",
    "opportunities",
    "opportunity",
)


def _theme_rgb(theme: int | None, tint: float | None = None) -> tuple[int, int, int] | None:
    # Approximate Office theme colors used by Delta scorecards.
    theme_map = {
        0: (255, 255, 255),
        1: (0, 0, 0),
        2: (238, 238, 238),
        3: (31, 73, 125),
        4: (79, 129, 189),
        5: (192, 80, 77),
        6: (155, 187, 89),
        7: (128, 100, 162),
        8: (75, 172, 198),
        9: (247, 150, 70),
    }
    if theme is None:
        return None
    base = theme_map.get(theme)
    if base is None:
        return (80, 80, 80)
    if not tint:
        return base
    r, g, b = base
    if tint < 0:
        factor = 1.0 + tint
        return (int(r * factor), int(g * factor), int(b * factor))
    return (
        int(r + (255 - r) * tint),
        int(g + (255 - g) * tint),
        int(b + (255 - b) * tint),
    )


def _color_rgb(color) -> tuple[int, int, int] | None:
    if color is None or color.type is None:
        return None
    if color.type == "rgb" and color.rgb:
        value = str(color.rgb)
        if value.startswith("urn:"):
            return None
        if len(value) == 8:
            value = value[2:]
        if len(value) == 6:
            try:
                return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]
            except ValueError:
                return None
    if color.type == "theme":
        return _theme_rgb(color.theme, getattr(color, "tint", None))
    if color.type == "indexed" and color.indexed is not None:
        idx = int(color.indexed)
        if 0 <= idx < len(COLOR_INDEX):
            value = COLOR_INDEX[idx]
            if isinstance(value, str) and len(value) >= 6:
                value = value[-6:]
                try:
                    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]
                except ValueError:
                    return None
    return None


def _fill_rgb(cell: Cell) -> tuple[int, int, int] | None:
    fill = cell.fill
    if fill is None or fill.fill_type in (None, "none"):
        return None
    return _color_rgb(fill.fgColor) or _color_rgb(fill.bgColor)


def _luminance(rgb: tuple[int, int, int]) -> float:
    r, g, b = rgb
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _is_dark(rgb: tuple[int, int, int] | None) -> bool:
    return rgb is not None and _luminance(rgb) < 55


def _cell_text(cell: Cell) -> str:
    value = cell.value
    if value is None:
        return ""
    return str(value).strip()


def _normalize_label(text: str) -> str:
    """Collapse NBSP / newlines / multi-space so TOTAL SCORE matches reliably."""
    if not text:
        return ""
    cleaned = (
        str(text)
        .replace("\xa0", " ")
        .replace("\u2007", " ")
        .replace("\u202f", " ")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )
    return " ".join(cleaned.split()).casefold()


def _is_total_score_label(text: str) -> bool:
    label = _normalize_label(text)
    if not label:
        return False
    compact = label.replace(" ", "")
    return (
        label == "total score"
        or compact == "totalscore"
        or label.startswith("total score")
        or "total score" in label
    )


def _row_has_content(ws: Worksheet, row: int, min_col: int, max_col: int) -> bool:
    for col in range(min_col, max_col + 1):
        if _cell_text(ws.cell(row, col)):
            return True
        if _fill_rgb(ws.cell(row, col)) is not None:
            return True
    return False


def _used_bounds(ws: Worksheet) -> tuple[int, int, int, int]:
    min_row, max_row = None, None
    min_col, max_col = None, None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None and _fill_rgb(cell) is None:
                continue
            r, c = cell.row, cell.column
            min_row = r if min_row is None else min(min_row, r)
            max_row = r if max_row is None else max(max_row, r)
            min_col = c if min_col is None else min(min_col, c)
            max_col = c if max_col is None else max(max_col, c)
    for merged in ws.merged_cells.ranges:
        min_row = merged.min_row if min_row is None else min(min_row, merged.min_row)
        max_row = merged.max_row if max_row is None else max(max_row, merged.max_row)
        min_col = merged.min_col if min_col is None else min(min_col, merged.min_col)
        max_col = merged.max_col if max_col is None else max(max_col, merged.max_col)
    if min_row is None:
        return 1, 1, 1, 1
    return min_row, max_row, min_col, max_col


def _normalize_month_token(text: str) -> str:
    cleaned = "".join(ch for ch in text.upper() if ch.isalpha())
    return cleaned


def _find_month_header_row(ws: Worksheet, min_row: int, max_row: int, min_col: int, max_col: int) -> int | None:
    """Locate the JAN..DEC / YTD / YE header row used by the System scorecard."""
    best_row = None
    best_hits = 0
    for row in range(min_row, min(max_row, min_row + 40) + 1):
        hits = 0
        for col in range(min_col, max_col + 1):
            token = _normalize_month_token(_cell_text(ws.cell(row, col)))
            if not token:
                continue
            if token in MONTH_HEADER_TOKENS or token[:3] in MONTH_HEADER_TOKENS:
                hits += 1
        if hits > best_hits:
            best_hits = hits
            best_row = row
    if best_hits >= 6:
        return best_row
    return None


def _header_band(ws: Worksheet, month_row: int, min_col: int, max_col: int, first_section_row: int) -> tuple[int, int]:
    """Include only the month header row (and an immediate label row), not the whole body."""
    start = month_row
    if month_row > 1 and _row_has_content(ws, month_row - 1, min_col, max_col):
        # Keep a short title row above months, but never more than 1 row.
        start = month_row - 1
    end = month_row
    # At most one WEIGHT/KPI label row directly under the months.
    probe = month_row + 1
    if probe < first_section_row:
        texts = " ".join(_cell_text(ws.cell(probe, c)).casefold() for c in range(min_col, min(min_col + 8, max_col + 1)))
        if "weight" in texts or texts.strip() == "kpi" or "total score" in texts:
            end = probe
    return start, end


def _is_category_title(text: str) -> bool:
    if not text or len(text) > 80:
        return False
    lower = text.casefold().strip()
    if lower in METRIC_LABELS or lower in KNOWN_KPI_NAMES:
        return False
    # Avoid matching bare "safety" inside KPI names like "Psych. Safety".
    if lower.startswith("psych"):
        return False
    token = _normalize_month_token(text)
    if token in MONTH_HEADER_TOKENS or token[:3] in MONTH_HEADER_TOKENS:
        return False
    if "(" in text and "%" in text:
        return True
    if any(k in lower for k in CATEGORY_KEYWORDS):
        return True
    return False


def _left_label_columns(min_col: int, max_col: int) -> range:
    # Real category rail is the far-left columns only (not the KPI name column).
    return range(min_col, min(min_col + 2, max_col + 1))


def _assign_default_titles(
    sections: list[tuple[int, int, str, bool]],
) -> list[tuple[int, int, str, bool]]:
    """Fill missing/generic titles using the known System category order."""
    out: list[tuple[int, int, str, bool]] = []
    for idx, (start, end, title, is_black) in enumerate(sections):
        cleaned = title.strip()
        generic = (
            not cleaned
            or cleaned.casefold().startswith("section ")
            or cleaned.casefold() in {"system", "system scorecard", "total score"}
        )
        if generic and idx < len(DEFAULT_SYSTEM_CATEGORIES):
            cleaned = DEFAULT_SYSTEM_CATEGORIES[idx]
        elif generic:
            cleaned = f"Section {idx + 1}"
        if idx == len(sections) - 1 and (
            is_black or "overall" in cleaned.casefold() or "opportunit" in cleaned.casefold()
        ):
            is_black = True
            if "overall" not in cleaned.casefold():
                cleaned = DEFAULT_SYSTEM_CATEGORIES[-1]
        out.append((start, end, cleaned, is_black))
    return out


def _vertical_category_merges(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[tuple[int, int, str, bool]]:
    """
    Category sections are tall merged cells in the leftmost columns, e.g.
    'Safety & Security (25.0%)' spanning many rows with a maroon/blue/orange fill.
    Ignores sheet-wide wrapper merges like a single 'System' label.
    """
    body_height = max(max_row - min_row + 1, 1)
    found: list[tuple[int, int, str, bool]] = []
    for merged in ws.merged_cells.ranges:
        height = merged.max_row - merged.min_row + 1
        width = merged.max_col - merged.min_col + 1
        if height < 4 or width > 2:
            continue
        if merged.min_col > min_col + 4:
            continue
        if merged.max_row < min_row or merged.min_row > max_row:
            continue
        # Skip wrappers that cover most of the sheet (seen as one 'System' block).
        if height >= max(20, int(body_height * 0.45)):
            continue
        title = _cell_text(ws.cell(merged.min_row, merged.min_col))
        if not title:
            continue
        if title.strip().casefold() in {"system", "system scorecard"} and height >= 20:
            continue
        rgb = _fill_rgb(ws.cell(merged.min_row, merged.min_col))
        is_black = _is_dark(rgb)
        found.append((merged.min_row, merged.max_row, title, is_black))

    found.sort(key=lambda item: item[0])
    return found


def _find_total_score_rows(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[int]:
    """Return 1-based rows that contain a TOTAL SCORE label anywhere in the used width."""
    starts: list[int] = []
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            if _is_total_score_label(_cell_text(ws.cell(row, col))):
                starts.append(row)
                break
    return starts


def _total_score_rows_via_excel_com(
    workbook_bytes: bytes,
    sheet_name: str,
    min_row: int,
    max_row: int,
) -> list[int]:
    """Read displayed cell text through Excel — catches labels openpyxl misses."""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return []

    excel = None
    wb = None
    starts: list[int] = []
    try:
        try:
            import pythoncom

            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None  # type: ignore
        with tempfile.TemporaryDirectory(prefix="mpr_scorecard_detect_") as tmp:
            path = Path(tmp) / "scorecards.xlsx"
            path.write_bytes(workbook_bytes)
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(str(path), ReadOnly=True, UpdateLinks=0)
            sheet = None
            for candidate in wb.Worksheets:
                if str(candidate.Name).strip().casefold() == sheet_name.casefold():
                    sheet = candidate
                    break
            if sheet is None:
                return []
            used = sheet.UsedRange
            values = used.Value
            if values is None:
                return []
            # UsedRange.Value is a tuple-of-tuples for multi-cell ranges.
            if not isinstance(values, tuple):
                values = ((values,),)
            elif values and not isinstance(values[0], tuple):
                values = (values,)
            origin_row = int(used.Row)
            for r_idx, row_vals in enumerate(values):
                row_num = origin_row + r_idx
                if row_num < min_row or row_num > max_row:
                    continue
                if not isinstance(row_vals, tuple):
                    row_vals = (row_vals,)
                for value in row_vals:
                    if value is None:
                        continue
                    if _is_total_score_label(str(value)):
                        starts.append(row_num)
                        break
            if starts:
                logger.info("Excel COM found TOTAL SCORE rows: %s", starts)
            return starts
    except Exception as exc:
        logger.info("Excel COM TOTAL SCORE scan unavailable: %s", exc)
        return []
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
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass


def _sections_from_total_score_starts(
    ws: Worksheet,
    starts: list[int],
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[tuple[int, int, str, bool]]:
    if len(starts) < 2:
        return []
    # Dedupe + sort; ignore TOTAL SCORE rows above the body.
    starts = sorted({row for row in starts if row >= min_row})
    if len(starts) < 2:
        return []

    sections: list[tuple[int, int, str, bool]] = []
    for idx, start in enumerate(starts):
        end = starts[idx + 1] - 1 if idx + 1 < len(starts) else max_row
        title = ""
        is_black = False
        for row in range(start, min(start + 10, end + 1)):
            for col in _left_label_columns(min_col, max_col):
                cell = ws.cell(row, col)
                text = _cell_text(cell)
                rgb = _fill_rgb(cell)
                if rgb is not None and _is_dark(rgb):
                    is_black = True
                if text and _is_category_title(text):
                    title = text
                    break
            if title:
                break
        if not title and idx < len(DEFAULT_SYSTEM_CATEGORIES):
            title = DEFAULT_SYSTEM_CATEGORIES[idx]
        elif not title:
            title = f"Section {idx + 1}"
        sections.append((start, end, title, is_black))
    return _assign_default_titles(sections)


def _total_score_section_starts(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    *,
    workbook_bytes: bytes | None = None,
    sheet_name: str = "System",
) -> list[tuple[int, int, str, bool]]:
    """
    Each category block in the live System tab starts with a grey TOTAL SCORE row.
    Use those rows as section boundaries.
    """
    starts = _find_total_score_rows(ws, min_row, max_row, min_col, max_col)
    if len(starts) < 2 and workbook_bytes:
        starts = _total_score_rows_via_excel_com(
            workbook_bytes, sheet_name, min_row, max_row
        )
    return _sections_from_total_score_starts(ws, starts, min_row, max_row, min_col, max_col)


def _default_category_bands(
    min_row: int,
    max_row: int,
) -> list[tuple[int, int, str, bool]]:
    """Last-resort split into the six known System categories by row bands."""
    names = DEFAULT_SYSTEM_CATEGORIES
    height = max_row - min_row + 1
    if height < len(names) * 4:
        return []
    sections: list[tuple[int, int, str, bool]] = []
    for idx, name in enumerate(names):
        start = min_row + (height * idx) // len(names)
        end = min_row + (height * (idx + 1)) // len(names) - 1
        if idx == len(names) - 1:
            end = max_row
        sections.append((start, max(start, end), name, idx == len(names) - 1))
    return sections


def _category_title_rows(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[tuple[int, int, str, bool]]:
    """Find category labels by text in the far-left rail only."""
    found: list[tuple[int, int, str, bool]] = []
    seen_rows: set[int] = set()
    for row in range(min_row, max_row + 1):
        for col in _left_label_columns(min_col, max_col):
            text = _cell_text(ws.cell(row, col))
            if not _is_category_title(text):
                continue
            if text.strip().casefold() in {"system", "system scorecard"}:
                continue
            if row in seen_rows:
                continue
            rgb = _fill_rgb(ws.cell(row, col))
            found.append((row, row, text, _is_dark(rgb)))
            seen_rows.add(row)
            break
    return _assign_default_titles(found) if len(found) >= 3 else []


def _fill_run_sections(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[tuple[int, int, str, bool]]:
    """Split on left-column fill-color changes across columns A-E."""
    best_col = min_col
    best_filled = -1
    for col in range(min_col, min(min_col + 5, max_col + 1)):
        filled = sum(1 for row in range(min_row, max_row + 1) if _fill_rgb(ws.cell(row, col)) is not None)
        if filled > best_filled:
            best_filled = filled
            best_col = col
    if best_filled < 8:
        return []

    sections: list[tuple[int, int, str, bool]] = []
    current_start = None
    current_rgb = None
    current_title = ""
    current_black = False

    def close(end_row: int) -> None:
        nonlocal current_start, current_rgb, current_title, current_black
        if current_start is None:
            return
        sections.append(
            (
                current_start,
                end_row,
                current_title or f"Section {len(sections) + 1}",
                current_black,
            )
        )
        current_start = None
        current_rgb = None
        current_title = ""
        current_black = False

    for row in range(min_row, max_row + 1):
        rgb = _fill_rgb(ws.cell(row, best_col))
        text = _cell_text(ws.cell(row, best_col))
        if rgb is None:
            continue
        if current_start is None:
            current_start = row
            current_rgb = rgb
            current_title = text if _is_category_title(text) else ""
            current_black = _is_dark(rgb)
            continue
        if rgb != current_rgb:
            close(row - 1)
            current_start = row
            current_rgb = rgb
            current_title = text if _is_category_title(text) else ""
            current_black = _is_dark(rgb)
        elif not current_title and _is_category_title(text):
            current_title = text
            current_black = current_black or _is_dark(rgb)

    if current_start is not None:
        close(max_row)

    kept = [s for s in sections if s[1] - s[0] >= 3]
    return _assign_default_titles(kept) if len(kept) >= 3 else []


def _choose_raw_sections(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    *,
    workbook_bytes: bytes | None = None,
    sheet_name: str = "System",
) -> list[tuple[int, int, str, bool]]:
    """Try several strategies; prefer TOTAL SCORE boundaries for the live workbook."""

    def total_score() -> list[tuple[int, int, str, bool]]:
        return _total_score_section_starts(
            ws,
            min_row,
            max_row,
            min_col,
            max_col,
            workbook_bytes=workbook_bytes,
            sheet_name=sheet_name,
        )

    candidates = [
        ("TOTAL SCORE rows", total_score),
        ("vertical merges", lambda: _vertical_category_merges(ws, min_row, max_row, min_col, max_col)),
        ("left fill runs", lambda: _fill_run_sections(ws, min_row, max_row, min_col, max_col)),
        ("category titles", lambda: _category_title_rows(ws, min_row, max_row, min_col, max_col)),
        ("default 6-band split", lambda: _default_category_bands(min_row, max_row)),
    ]
    for name, factory in candidates:
        found = factory()
        # Prefer a full System scorecard (5–6 categories). Accept 3+ as usable.
        if len(found) >= 3:
            if name == "default 6-band split":
                logger.warning(
                    "Using even 6-category row split R%s-%s (TOTAL SCORE labels not readable)",
                    min_row,
                    max_row,
                )
            logger.info("System section strategy=%s count=%s", name, len(found))
            return _assign_default_titles(found)
        logger.info("System section strategy=%s count=%s (skip)", name, len(found))
    return []


def detect_system_layout(
    workbook_bytes: bytes,
    sheet_name: str = "System",
) -> SystemLayout:
    """
    Parse the System scorecard layout matching the live GSE file:
    month header (JAN..YE) + vertical category bars on the left.
    """
    # data_only=False keeps labels/formulas visible for section detection.
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    if sheet_name not in wb.sheetnames:
        match = next((n for n in wb.sheetnames if n.strip().casefold() == sheet_name.casefold()), None)
        if match is None:
            raise ValueError(f"Sheet {sheet_name!r} not found. Available: {wb.sheetnames}")
        sheet_name = match
    ws = wb[sheet_name]
    min_row, max_row, min_col, max_col = _used_bounds(ws)

    month_row = _find_month_header_row(ws, min_row, max_row, min_col, max_col)
    body_start = (month_row + 1) if month_row else min_row

    raw_sections = _choose_raw_sections(
        ws,
        body_start,
        max_row,
        min_col,
        max_col,
        workbook_bytes=workbook_bytes,
        sheet_name=sheet_name,
    )
    if not raw_sections:
        logger.warning(
            "Could not split System categories (cols %s-%s rows %s-%s); capturing full body",
            get_column_letter(min_col),
            get_column_letter(max_col),
            body_start,
            max_row,
        )
        raw_sections = [(body_start, max_row, "System", False)]

    # Extend each section to the row before the next category starts.
    normalized: list[tuple[int, int, str, bool]] = []
    for idx, (start, end, title, is_black) in enumerate(raw_sections):
        next_start = raw_sections[idx + 1][0] if idx + 1 < len(raw_sections) else max_row + 1
        end = max(end, next_start - 1)
        while end > start and not _row_has_content(ws, end, min_col, max_col):
            end -= 1
        # Mark trailing dark totals block as black when fill says so.
        if not is_black:
            for col in _left_label_columns(min_col, max_col):
                if _is_dark(_fill_rgb(ws.cell(start, col))):
                    is_black = True
                    break
        normalized.append((start, end, title, is_black))

    first_section_row = normalized[0][0]
    if month_row:
        header_start, header_end = _header_band(ws, month_row, min_col, max_col, first_section_row)
    else:
        header_start = min_row
        header_end = max(min_row, first_section_row - 1)

    sections = [
        Section(
            index=idx,
            title=title,
            start_row=start,
            end_row=end,
            start_col=min_col,
            end_col=max_col,
            is_black=is_black,
        )
        for idx, (start, end, title, is_black) in enumerate(normalized)
    ]

    layout = SystemLayout(
        header_start_row=header_start,
        header_end_row=header_end,
        start_col=min_col,
        end_col=max_col,
        sections=sections,
    )
    wb.close()
    logger.info(
        "System layout: header rows %s-%s cols %s-%s; sections=%s",
        layout.header_start_row,
        layout.header_end_row,
        get_column_letter(layout.start_col),
        get_column_letter(layout.end_col),
        [(s.index, s.title, f"R{s.start_row}-{s.end_row}", s.is_black) for s in sections],
    )
    return layout


def detect_system_sections(
    workbook_bytes: bytes,
    sheet_name: str = "System",
) -> list[Section]:
    """Backward-compatible helper returning only category sections."""
    return detect_system_layout(workbook_bytes, sheet_name=sheet_name).sections


def select_sections_for_slide(
    sections: list[Section],
    *,
    mode: str,
    count: int = 3,
    include_black: bool = False,
    match: list[str] | None = None,
    indices: list[int] | None = None,
) -> list[Section]:
    """Choose which sections belong on a System Scorecard slide."""
    if not sections:
        return []

    if mode == "indices":
        chosen = [sections[i] for i in (indices or []) if 0 <= i < len(sections)]
        return sorted({s.index: s for s in chosen}.values(), key=lambda s: s.start_row)

    if mode == "match":
        patterns = [p.strip() for p in (match or []) if str(p).strip()]
        if not patterns:
            raise ValueError("sections: match requires a non-empty match list")
        chosen: list[Section] = []
        used: set[int] = set()
        for pattern in patterns:
            needle = pattern.casefold()
            for section in sections:
                if section.index in used:
                    continue
                if needle in section.title.casefold():
                    chosen.append(section)
                    used.add(section.index)
                    break
        chosen = sorted({s.index: s for s in chosen}.values(), key=lambda s: s.start_row)
        if chosen:
            return chosen
        if indices:
            logger.warning(
                "Name match failed for %s; falling back to indices %s (available=%s)",
                patterns,
                indices,
                [s.title for s in sections],
            )
            fallback = select_sections_for_slide(sections, mode="indices", indices=indices)
            if fallback:
                return fallback
            # Indices out of range (e.g. only one mega-section) — take first/last third.
            if len(sections) == 1:
                return list(sections)
            mid = max(1, len(sections) // 2)
            if min(indices) >= mid:
                return sections[mid:]
            return sections[:mid]
        logger.warning(
            "No System sections matched %s. Available: %s",
            patterns,
            [s.title for s in sections],
        )
        return chosen

    black = [s for s in sections if s.is_black]
    non_black = [s for s in sections if not s.is_black] or list(sections)

    if mode == "first":
        chosen = non_black[:count]
    elif mode == "last":
        chosen = non_black[-count:] if len(non_black) >= count else list(non_black)
        if include_black:
            for b in black:
                if b not in chosen:
                    chosen.append(b)
            if not black and sections[-1] not in chosen:
                chosen.append(sections[-1])
    else:
        raise ValueError(f"Unknown section mode: {mode!r}")

    chosen = sorted(chosen, key=lambda s: s.start_row)
    return chosen


_LAYOUT_CACHE: dict[tuple[int, str], SystemLayout] = {}


def _cached_system_layout(workbook_bytes: bytes, sheet_name: str) -> SystemLayout:
    key = (hash(workbook_bytes), sheet_name.casefold())
    cached = _LAYOUT_CACHE.get(key)
    if cached is not None:
        return cached
    layout = detect_system_layout(workbook_bytes, sheet_name=sheet_name)
    _LAYOUT_CACHE[key] = layout
    return layout


def _range_address(start_row: int, end_row: int, start_col: int, end_col: int) -> str:
    return f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"


def _capture_via_excel_com(workbook_path: Path, sheet_name: str, range_addr: str) -> bytes:
    """True Excel screenshot via CopyPicture — matches template look when pywin32 is installed."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32\n"
            "Then re-run main.py for template-quality System scorecard screenshots."
        ) from exc
    from PIL import ImageGrab

    excel = None
    wb = None
    try:
        try:
            excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = excel.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=True, UpdateLinks=0)
        ws = None
        for sheet in wb.Worksheets:
            if str(sheet.Name).strip().casefold() == sheet_name.casefold():
                ws = sheet
                break
        if ws is None:
            raise ValueError(f"Worksheet {sheet_name!r} not found in Excel COM open")
        ws.Activate()
        try:
            excel.ActiveWindow.Zoom = 100
        except Exception:
            pass
        rng = ws.Range(range_addr)
        rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
        img = None
        for _ in range(30):
            time.sleep(0.2)
            img = ImageGrab.grabclipboard()
            if img is not None:
                break
        if img is None:
            # Fallback: Copy then grab
            rng.Copy()
            time.sleep(0.5)
            img = ImageGrab.grabclipboard()
        if img is None:
            raise RuntimeError(f"Excel clipboard capture failed for {sheet_name}!{range_addr}")
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
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


def _font(size: int = 12):
    for name in (
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ):
        path = Path(name)
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size=size)
            except OSError:
                continue
    return ImageFont.load_default()


def _capture_via_render(
    workbook_bytes: bytes,
    sheet_name: str,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> bytes:
    """Approximate Excel look when Windows Excel COM is unavailable."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=True)
    match = next((n for n in wb.sheetnames if n.strip().casefold() == sheet_name.casefold()), sheet_name)
    ws = wb[match]

    col_widths = []
    for c in range(start_col, end_col + 1):
        letter = get_column_letter(c)
        width = ws.column_dimensions[letter].width
        col_widths.append(int(max(float(width or 10.0), 4.0) * 9))
    row_heights = []
    for r in range(start_row, end_row + 1):
        height = ws.row_dimensions[r].height
        row_heights.append(int(max(float(height or 15.0), 12.0) * 1.35))

    img_w = max(sum(col_widths), 40)
    img_h = max(sum(row_heights), 40)
    image = Image.new("RGB", (img_w, img_h), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    font = _font(11)
    bold = _font(12)

    y = 0
    for r_idx, r in enumerate(range(start_row, end_row + 1)):
        x = 0
        rh = row_heights[r_idx]
        for c_idx, c in enumerate(range(start_col, end_col + 1)):
            cw = col_widths[c_idx]
            cell = ws.cell(r, c)
            fill = _fill_rgb(cell) or (255, 255, 255)
            draw.rectangle([x, y, x + cw, y + rh], fill=fill, outline=(200, 200, 200))
            text = _cell_text(cell)
            if text:
                text_rgb = _color_rgb(cell.font.color) if cell.font and cell.font.color else None
                if text_rgb is None:
                    text_rgb = (255, 255, 255) if _is_dark(fill) else (20, 20, 20)
                use_font = bold if (cell.font and cell.font.bold) else font
                draw.text((x + 4, y + max(2, (rh - 14) // 2)), text[:60], fill=text_rgb, font=use_font)
            x += cw
        y += rh

    wb.close()
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return buf.getvalue()


def capture_range_png(
    workbook_bytes: bytes,
    sheet_name: str,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    *,
    prefer_excel_com: bool = True,
) -> bytes:
    """Screenshot (or render) an Excel A1-style range to PNG bytes."""
    range_addr = _range_address(start_row, end_row, start_col, end_col)
    if prefer_excel_com:
        try:
            with tempfile.TemporaryDirectory(prefix="mpr_scorecard_") as tmp:
                path = Path(tmp) / "scorecards.xlsx"
                path.write_bytes(workbook_bytes)
                png = _capture_via_excel_com(path, sheet_name, range_addr)
                logger.info("Captured %s!%s via Excel COM (%s bytes)", sheet_name, range_addr, len(png))
                return png
        except Exception as exc:
            logger.warning(
                "Excel COM capture unavailable (%s). "
                "For true Excel screenshots on Windows run: pip install pywin32. "
                "Using Pillow render fallback for now.",
                exc,
            )

    png = _capture_via_render(workbook_bytes, sheet_name, start_row, end_row, start_col, end_col)
    logger.info("Rendered %s!%s via Pillow (%s bytes)", sheet_name, range_addr, len(png))
    return png


def _stitch_vertical_many(png_parts: list[bytes]) -> bytes:
    images = [Image.open(io.BytesIO(p)).convert("RGB") for p in png_parts if p]
    if not images:
        raise ValueError("No images to stitch")
    width = max(im.width for im in images)
    height = sum(im.height for im in images)
    canvas = Image.new("RGB", (width, height), (255, 255, 255))
    y = 0
    for im in images:
        canvas.paste(im, (0, y))
        y += im.height
    buf = io.BytesIO()
    canvas.save(buf, format="PNG")
    return buf.getvalue()


def _sections_are_contiguous(sections: list[Section]) -> bool:
    ordered = sorted(sections, key=lambda s: s.start_row)
    for idx in range(1, len(ordered)):
        if ordered[idx].start_row > ordered[idx - 1].end_row + 2:
            return False
    return True


def capture_sections_png(
    workbook_bytes: bytes,
    sheet_name: str,
    sections: list[Section],
    *,
    layout: SystemLayout | None = None,
    prefer_excel_com: bool = True,
) -> bytes | None:
    """
    Screenshot the System scorecard like the live Excel / template view:
    month header (JAN..YE) + selected vertical category blocks.
    """
    if not sections:
        return None
    if layout is None:
        layout = detect_system_layout(workbook_bytes, sheet_name=sheet_name)

    ordered = sorted(sections, key=lambda s: s.start_row)
    start_col, end_col = layout.start_col, layout.end_col

    def _cap(r1: int, r2: int) -> bytes:
        return capture_range_png(
            workbook_bytes,
            sheet_name,
            r1,
            r2,
            start_col,
            end_col,
            prefer_excel_com=prefer_excel_com,
        )

    # Contiguous first-N categories: one CopyPicture range (template look).
    if _sections_are_contiguous(ordered) and not layout.needs_header_stitch(ordered):
        return _cap(layout.header_start_row, ordered[-1].end_row)

    parts: list[bytes] = [_cap(layout.header_start_row, layout.header_end_row)]
    if _sections_are_contiguous(ordered):
        parts.append(_cap(ordered[0].start_row, ordered[-1].end_row))
    else:
        # Non-adjacent categories (e.g. Finance + People with a gap): capture each block.
        for section in ordered:
            parts.append(_cap(section.start_row, section.end_row))
    return _stitch_vertical_many(parts)


def place_picture_on_slide(
    slide,
    png_bytes: bytes,
    *,
    left: int = DEFAULT_LEFT,
    top: int = DEFAULT_TOP,
    max_width: int = DEFAULT_WIDTH,
    max_height: int = DEFAULT_HEIGHT,
    fit: str = "fill",
) -> None:
    """Place a PNG into the content area.

    fit=fill  -> stretch exactly into the box (matches template picture slots)
    fit=contain -> preserve aspect ratio inside the box
    """
    from pptx.util import Emu

    with Image.open(io.BytesIO(png_bytes)) as img:
        img_w, img_h = img.size
    if img_w <= 0 or img_h <= 0:
        return

    if fit == "contain":
        scale = min(max_width / img_w, max_height / img_h)
        width = int(img_w * scale)
        height = int(img_h * scale)
        left_pos = left + max(0, (max_width - width) // 2)
        top_pos = top + max(0, (max_height - height) // 2)
    else:
        # Fill the template picture rectangle exactly.
        width = int(max_width)
        height = int(max_height)
        left_pos = int(left)
        top_pos = int(top)

    slide.shapes.add_picture(
        io.BytesIO(png_bytes),
        Emu(left_pos),
        Emu(top_pos),
        Emu(width),
        Emu(height),
    )


def _largest_picture(slide):
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    pictures = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if not pictures:
        return None
    return max(pictures, key=lambda s: int(s.width) * int(s.height))


def _remove_slide_pictures(slide) -> int:
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    removed = 0
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            element = shape._element
            element.getparent().remove(element)
            removed += 1
    return removed


def resolve_sheet_name(
    workbook_bytes: bytes,
    *,
    sheet: str | None = None,
    sheet_index: int | None = None,
    sheet_match: list[str] | None = None,
    sheet_match_index: int = 0,
    available: list[str] | None = None,
) -> str:
    """Resolve a worksheet name from explicit name, index, or substring match."""
    names = available
    if names is None:
        wb = load_workbook(io.BytesIO(workbook_bytes), read_only=True, data_only=False)
        try:
            names = list(wb.sheetnames)
        finally:
            wb.close()

    if sheet:
        if sheet in names:
            return sheet
        exact = next((n for n in names if n.strip().casefold() == sheet.casefold()), None)
        if exact:
            return exact

    patterns = [str(p).strip().casefold() for p in (sheet_match or []) if str(p).strip()]
    if patterns:
        hits = [n for n in names if all(p in n.casefold() for p in patterns)]
        if not hits:
            hits = [n for n in names if any(p in n.casefold() for p in patterns)]
        if hits:
            idx = int(sheet_match_index or 0)
            if idx < 0:
                idx = len(hits) + idx
            if 0 <= idx < len(hits):
                return hits[idx]
            # Match found but index out of range — fall through to sheet_index.

    if sheet_index is not None:
        idx = int(sheet_index)
        if 0 <= idx < len(names):
            return names[idx]
        raise ValueError(f"sheet_index={idx} out of range for sheets {names}")

    if sheet:
        raise ValueError(f"Sheet {sheet!r} not found. Available: {names}")
    raise ValueError(f"Could not resolve sheet from index/match. Available: {names}")


def capture_sheet_png(
    workbook_bytes: bytes,
    sheet_name: str,
    *,
    prefer_excel_com: bool = True,
    max_rows: int | None = 120,
    max_cols: int | None = 40,
) -> bytes:
    """Screenshot the used range of any worksheet (entity / workings / System)."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    match = next((n for n in wb.sheetnames if n.strip().casefold() == sheet_name.casefold()), None)
    if match is None:
        wb.close()
        raise ValueError(f"Sheet {sheet_name!r} not found. Available: {wb.sheetnames}")
    ws = wb[match]
    min_row, max_row, min_col, max_col = _used_bounds(ws)
    wb.close()

    if max_rows is not None:
        max_row = min(max_row, min_row + int(max_rows) - 1)
    if max_cols is not None:
        max_col = min(max_col, min_col + int(max_cols) - 1)

    logger.info(
        "Capturing full sheet %s!%s",
        match,
        _range_address(min_row, max_row, min_col, max_col),
    )
    return capture_range_png(
        workbook_bytes,
        match,
        min_row,
        max_row,
        min_col,
        max_col,
        prefer_excel_com=prefer_excel_com,
    )


def apply_scorecard_screenshot(slide, data, element: dict) -> bool:
    """Replace a template picture with a live Excel screenshot (System sections or full sheet)."""
    workbook = element.get("workbook", "scorecards")
    prefer_com = bool(element.get("prefer_excel_com", True))
    replace_existing = bool(element.get("replace_existing_pictures", True))
    fit = str(element.get("fit", "fill")).lower()
    mode = str(element.get("sections", element.get("capture", "auto"))).lower()

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("Scorecard screenshot skipped: %s", exc)
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
            sheet_match=element.get("sheet_match"),
            sheet_match_index=int(element.get("sheet_match_index", 0) or 0),
            available=available or None,
        )
    except Exception as exc:
        logger.warning("Scorecard sheet resolve failed: %s", exc)
        return False

    explicit = element.get("range")
    png: bytes | None = None

    # Full-sheet / explicit-range capture for entity tabs, workings, etc.
    wants_full_sheet = mode in {"sheet", "full", "all", "used_range"} or (
        mode == "auto" and sheet_name.casefold() != "system"
    )
    if explicit:
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(str(explicit))
        png = capture_range_png(
            workbook_bytes,
            sheet_name,
            min_row,
            max_row,
            min_col,
            max_col,
            prefer_excel_com=prefer_com,
        )
    elif wants_full_sheet:
        try:
            png = capture_sheet_png(
                workbook_bytes,
                sheet_name,
                prefer_excel_com=prefer_com,
                max_rows=element.get("max_rows", 120),
                max_cols=element.get("max_cols", 40),
            )
        except Exception as exc:
            logger.warning("Full-sheet capture failed for %s: %s", sheet_name, exc)
            return False
    else:
        # System tab: split into named category blocks.
        try:
            layout = _cached_system_layout(workbook_bytes, sheet_name)
        except Exception as exc:
            logger.warning("Could not detect System layout: %s", exc)
            return False

        count = int(element.get("count", 3))
        include_black = bool(element.get("include_black", False))
        match = element.get("match") or []
        fallback_indices = element.get("fallback_indices") or element.get("indices") or []
        section_mode = mode if mode in {"match", "indices", "first", "last"} else (
            "match" if match else "first"
        )
        chosen = select_sections_for_slide(
            layout.sections,
            mode=section_mode,
            count=count,
            include_black=include_black,
            match=list(match) if match else None,
            indices=[int(i) for i in fallback_indices] if fallback_indices else None,
        )
        if not chosen:
            logger.warning(
                "No System sections selected for mode=%s match=%s available=%s",
                section_mode,
                match,
                [s.title for s in layout.sections],
            )
            return False
        png = capture_sections_png(
            workbook_bytes,
            sheet_name,
            chosen,
            layout=layout,
            prefer_excel_com=prefer_com,
        )
        logger.info("System screenshot sections=%s", [s.title for s in chosen])

    if not png:
        return False

    # Use the template picture slot geometry when present so placement matches the deck.
    target = _largest_picture(slide)
    if target is not None:
        left, top, width, height = int(target.left), int(target.top), int(target.width), int(target.height)
    else:
        left = int(element.get("left", DEFAULT_LEFT))
        top = int(element.get("top", DEFAULT_TOP))
        width = int(element.get("max_width", DEFAULT_WIDTH))
        height = int(element.get("max_height", DEFAULT_HEIGHT))

    if replace_existing:
        removed = _remove_slide_pictures(slide)
        if removed:
            logger.info("Removed %s template picture(s) before placing live scorecard image", removed)

    place_picture_on_slide(
        slide,
        png,
        left=left,
        top=top,
        max_width=width,
        max_height=height,
        fit=fit,
    )
    logger.info(
        "Placed scorecard image from %s!%s box=(%s,%s,%s,%s) fit=%s",
        workbook,
        sheet_name,
        left,
        top,
        width,
        height,
        fit,
    )
    return True
