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
    """One visual block on the System scorecard sheet."""

    index: int
    title: str
    start_row: int  # 1-based inclusive
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


def _row_fill_ratio(ws: Worksheet, row: int, min_col: int, max_col: int) -> tuple[float, int, int]:
    """Return (dark_ratio_among_filled, filled_count, dark_count)."""
    filled = 0
    dark = 0
    for col in range(min_col, max_col + 1):
        rgb = _fill_rgb(ws.cell(row, col))
        if rgb is None:
            continue
        filled += 1
        if _is_dark(rgb):
            dark += 1
    if filled == 0:
        return 0.0, 0, 0
    return dark / filled, filled, dark


def _row_is_section_header(ws: Worksheet, row: int, min_col: int, max_col: int) -> bool:
    """Section banners: colored bar across many columns, often merged, with a short title."""
    width = max_col - min_col + 1
    dark_ratio, filled, dark = _row_fill_ratio(ws, row, min_col, max_col)
    fill_coverage = filled / max(width, 1)

    title = ""
    title_col = min_col
    for col in range(min_col, max_col + 1):
        title = _cell_text(ws.cell(row, col))
        if title:
            title_col = col
            break
    if not title or len(title) > 80:
        return False

    non_empty = sum(1 for c in range(min_col, max_col + 1) if _cell_text(ws.cell(row, c)))
    merged_banner = False
    for merged in ws.merged_cells.ranges:
        if merged.min_row == row == merged.max_row and merged.max_col - merged.min_col >= 2:
            if merged.min_col <= title_col <= merged.max_col:
                merged_banner = True
                break

    # Strong signal: wide colored banner with few text cells.
    if fill_coverage >= 0.5 and non_empty <= 3:
        return True
    if merged_banner and (fill_coverage >= 0.2 or dark > 0):
        return True
    # Dark totals / black strip
    if dark_ratio >= 0.5 and fill_coverage >= 0.4 and non_empty <= 3:
        return True
    return False


def detect_system_sections(
    workbook_bytes: bytes,
    sheet_name: str = "System",
) -> list[Section]:
    """Split the System sheet into visual sections (tables / banners)."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        match = next((n for n in wb.sheetnames if n.strip().casefold() == sheet_name.casefold()), None)
        if match is None:
            raise ValueError(f"Sheet {sheet_name!r} not found. Available: {wb.sheetnames}")
        sheet_name = match
    ws = wb[sheet_name]
    min_row, max_row, min_col, max_col = _used_bounds(ws)

    header_rows = [
        r for r in range(min_row, max_row + 1) if _row_is_section_header(ws, r, min_col, max_col)
    ]

    # Fallback: split on blank-row gaps when banners are not styled.
    if len(header_rows) < 2:
        header_rows = []
        blank_run = 0
        started = False
        for r in range(min_row, max_row + 1):
            if _row_has_content(ws, r, min_col, max_col):
                if not started or blank_run >= 2:
                    header_rows.append(r)
                    started = True
                blank_run = 0
            else:
                blank_run += 1
        if not header_rows:
            header_rows = [min_row]

    header_rows = sorted(set(header_rows))
    sections: list[Section] = []
    for idx, start in enumerate(header_rows):
        end = header_rows[idx + 1] - 1 if idx + 1 < len(header_rows) else max_row
        while end > start and not _row_has_content(ws, end, min_col, max_col):
            end -= 1
        title = ""
        for col in range(min_col, max_col + 1):
            title = _cell_text(ws.cell(start, col))
            if title:
                break

        # Black section = dark banner row (or mostly dark fills in the first rows).
        dark_ratio, filled, _dark = _row_fill_ratio(ws, start, min_col, max_col)
        is_black = filled > 0 and dark_ratio >= 0.45
        if not is_black:
            dark_cells = 0
            checked = 0
            for r in range(start, min(start + 2, end + 1)):
                for c in range(min_col, max_col + 1):
                    rgb = _fill_rgb(ws.cell(r, c))
                    if rgb is None:
                        continue
                    checked += 1
                    if _is_dark(rgb):
                        dark_cells += 1
            is_black = checked > 0 and dark_cells / checked >= 0.45

        sections.append(
            Section(
                index=idx,
                title=title or f"Section {idx + 1}",
                start_row=start,
                end_row=end,
                start_col=min_col,
                end_col=max_col,
                is_black=is_black,
            )
        )

    wb.close()
    logger.info(
        "Detected %s System sections: %s",
        len(sections),
        [(s.index, s.title, s.range_address, s.is_black) for s in sections],
    )
    return sections


def select_sections_for_slide(
    sections: list[Section],
    *,
    mode: str,
    count: int = 3,
    include_black: bool = False,
) -> list[Section]:
    """Choose which sections belong on a System Scorecard slide."""
    if not sections:
        return []

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
            # If no color-flagged black section, append the true last sheet section
            # when it was excluded (often a dark totals strip).
            if not black and sections[-1] not in chosen:
                chosen.append(sections[-1])
    else:
        raise ValueError(f"Unknown section mode: {mode!r}")

    chosen = sorted(chosen, key=lambda s: s.start_row)
    return chosen


def _combine_range(sections: list[Section]) -> tuple[int, int, int, int]:
    return (
        min(s.start_row for s in sections),
        max(s.end_row for s in sections),
        min(s.start_col for s in sections),
        max(s.end_col for s in sections),
    )


def _range_address(start_row: int, end_row: int, start_col: int, end_col: int) -> str:
    return f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"


def _capture_via_excel_com(workbook_path: Path, sheet_name: str, range_addr: str) -> bytes:
    import win32com.client  # type: ignore
    from PIL import ImageGrab

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=True, UpdateLinks=0)
        ws = None
        for sheet in wb.Worksheets:
            if str(sheet.Name).strip().casefold() == sheet_name.casefold():
                ws = sheet
                break
        if ws is None:
            raise ValueError(f"Worksheet {sheet_name!r} not found in Excel COM open")
        rng = ws.Range(range_addr)
        # Appearance=1 (xlScreen), Format=2 (xlBitmap)
        rng.CopyPicture(Appearance=1, Format=2)
        img = None
        for _ in range(20):
            time.sleep(0.15)
            img = ImageGrab.grabclipboard()
            if img is not None:
                break
        if img is None:
            raise RuntimeError(f"Excel clipboard capture failed for {sheet_name}!{range_addr}")
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


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
                use_font = bold if (cell.font and cell.font.bold) or _row_is_section_header(ws, r, start_col, end_col) else font
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
            logger.warning("Excel COM capture unavailable (%s); using rendered fallback", exc)

    png = _capture_via_render(workbook_bytes, sheet_name, start_row, end_row, start_col, end_col)
    logger.info("Rendered %s!%s via Pillow (%s bytes)", sheet_name, range_addr, len(png))
    return png


def capture_sections_png(
    workbook_bytes: bytes,
    sheet_name: str,
    sections: list[Section],
    *,
    prefer_excel_com: bool = True,
) -> bytes | None:
    if not sections:
        return None
    start_row, end_row, start_col, end_col = _combine_range(sections)
    return capture_range_png(
        workbook_bytes,
        sheet_name,
        start_row,
        end_row,
        start_col,
        end_col,
        prefer_excel_com=prefer_excel_com,
    )


def place_picture_on_slide(
    slide,
    png_bytes: bytes,
    *,
    left: int = DEFAULT_LEFT,
    top: int = DEFAULT_TOP,
    max_width: int = DEFAULT_WIDTH,
    max_height: int = DEFAULT_HEIGHT,
) -> None:
    """Add a PNG centered in the content area, preserving aspect ratio."""
    from pptx.util import Emu

    with Image.open(io.BytesIO(png_bytes)) as img:
        img_w, img_h = img.size
    if img_w <= 0 or img_h <= 0:
        return

    scale = min(max_width / img_w, max_height / img_h)
    width = int(img_w * scale)
    height = int(img_h * scale)
    left_pos = left + max(0, (max_width - width) // 2)
    top_pos = top + max(0, (max_height - height) // 2)
    slide.shapes.add_picture(io.BytesIO(png_bytes), Emu(left_pos), Emu(top_pos), Emu(width), Emu(height))


def apply_scorecard_screenshot(slide, data, element: dict) -> bool:
    """Element handler: screenshot System scorecard sections onto a slide."""
    workbook = element.get("workbook", "scorecards")
    sheet_name = element.get("sheet", "System")
    mode = element.get("sections", "first")
    count = int(element.get("count", 3))
    include_black = bool(element.get("include_black", False))
    prefer_com = bool(element.get("prefer_excel_com", True))

    try:
        workbook_bytes = data.store.workbook_bytes(workbook)
    except FileNotFoundError as exc:
        logger.warning("Scorecard screenshot skipped: %s", exc)
        return False

    try:
        sections = detect_system_sections(workbook_bytes, sheet_name=sheet_name)
    except Exception as exc:
        logger.warning("Could not detect System sections: %s", exc)
        return False

    chosen = select_sections_for_slide(
        sections,
        mode=mode,
        count=count,
        include_black=include_black,
    )
    if not chosen:
        logger.warning("No System sections selected for mode=%s count=%s", mode, count)
        return False

    # Optional explicit A1 ranges override auto selection.
    explicit = element.get("range")
    if explicit:
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(explicit)
        png = capture_range_png(
            workbook_bytes,
            sheet_name,
            min_row,
            max_row,
            min_col,
            max_col,
            prefer_excel_com=prefer_com,
        )
    else:
        png = capture_sections_png(
            workbook_bytes,
            sheet_name,
            chosen,
            prefer_excel_com=prefer_com,
        )
    if not png:
        return False

    place_picture_on_slide(
        slide,
        png,
        left=int(element.get("left", DEFAULT_LEFT)),
        top=int(element.get("top", DEFAULT_TOP)),
        max_width=int(element.get("max_width", DEFAULT_WIDTH)),
        max_height=int(element.get("max_height", DEFAULT_HEIGHT)),
    )
    logger.info(
        "Placed System scorecard image on slide (%s section(s): %s)",
        len(chosen),
        [s.title for s in chosen],
    )
    return True
