"""Fill the GSE MPR PowerPoint template with Excel data."""

from __future__ import annotations

import io
import logging
import math
import re
from datetime import date
from pathlib import Path

import pandas as pd
import yaml
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

from data_lookup import lookup_kpi_exhaustive, lookup_monthly_series_exhaustive, lookup_ytd_exhaustive
from mpr_data import MONTH_LABELS, MprData
from ppt_format import (
    clear_table_data_rows,
    clear_text_frame_content,
    set_cell_text_preserve,
    set_paragraph_text_preserve,
    set_text_frame_preserve,
)
from ppt_missing import (
    EMPTY_CELL,
    chart_insert_message,
    data_not_found,
    is_missing_text,
    is_placeholder_text,
)
from scorecard_data import (
    build_scorecard_row_index,
    contextual_scorecard_key,
    find_scorecard_header_row,
    find_scorecard_table_header_row,
    format_scorecard_cell,
    load_system_scorecard,
    scorecard_column_map,
)
from scorecard_layout import fit_scorecard_table
from gir_slide import fill_gir_tables, load_gir_dataset
from narrative_boxes import clear_manual_narrative_boxes
from people_slide import chart_configs_from_element, fill_people_table
from picture_replace import (
    fill_styled_scorecard_table,
    fill_styled_worksheet_table,
    fill_table_from_dataframe,
    paste_dataframe_at_placeholder,
    paste_dataframe_on_slide,
    remove_data_placeholders,
)
from safety_compliance import fill_compliance_slide
from scorecard_style import apply_compliance_table_styles, apply_scorecard_block_styles

logger = logging.getLogger(__name__)

CHART_CATEGORIES = MONTH_LABELS + ["YTD"]

DEFAULT_AGENDA_SECTIONS = [
    {"title": "Safety", "time": "10 min"},
    {"title": "People", "time": "10 min"},
    {"title": "Finance", "time": "10 min"},
    {"title": "Customer Experience", "time": "10 min"},
    {"title": "North (M)", "time": "10 min"},
    {"title": "South (M)", "time": "10 min"},
    {"title": "Stationary", "time": "15 min"},
    {"title": "Galley", "time": "10 min"},
    {"title": "Closing", "time": "3 min"},
]

DEFAULT_AGENDA_HEADER = {"title": "GSE Monthly Performance Summary", "time": "5 min"}

ACTUALS_SOURCE = "MPR Actuals and Goals_v2.xlsx"
SCORECARDS_SOURCE = "2026 - GSE Scorecards.xlsx"


def _safe_number(value: float | None) -> float | None:
    if value is None:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number


def _fmt(value: float | None, decimals: int = 2) -> str:
    number = _safe_number(value)
    if number is None:
        return ""
    if decimals == 0:
        return f"{round(number):.0f}"
    return f"{number:.{decimals}f}"


def _fmt_diff(actual: float | None, goal: float | None, decimals: int = 2) -> str:
    actual_n = _safe_number(actual)
    goal_n = _safe_number(goal)
    if actual_n is None or goal_n is None:
        return ""
    diff = actual_n - goal_n
    if diff < 0:
        return f"({abs(diff):.{decimals}f})"
    return _fmt(diff, decimals)


def _fmt_cell(value: float | None, *, decimals: int = 2, empty: str = EMPTY_CELL) -> str:
    out = _fmt(value, decimals)
    return out if out else empty


def _fmt_diff_cell(
    actual: float | None,
    goal: float | None,
    *,
    decimals: int = 2,
    empty: str = EMPTY_CELL,
) -> str:
    out = _fmt_diff(actual, goal, decimals)
    return out if out else empty


def _kpi_patterns_from_config(config: dict, key: str, *extra: str) -> list[str]:
    val = config.get("kpi_mappings", {}).get(key, key)
    patterns = [str(item) for item in val] if isinstance(val, list) else [str(val)]
    for item in extra:
        if item not in patterns:
            patterns.append(item)
    return patterns


def _fmt_cell_or_data_not_found(value: float | None, what: str, *, decimals: int = 2) -> str:
    number = _safe_number(value)
    if number is None:
        return data_not_found(what)
    return _fmt(number, decimals)


def _fmt_diff_or_data_not_found(
    actual: float | None,
    goal: float | None,
    what: str,
    *,
    decimals: int = 2,
) -> str:
    if _safe_number(actual) is None or _safe_number(goal) is None:
        return data_not_found(what)
    return _fmt_diff(actual, goal, decimals)


def _replace_shape_with_notice(slide, shape, message: str) -> None:
    """Replace a chart/picture shape with a short notice textbox."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    _remove_shape(shape)
    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.word_wrap = True
    set_text_frame_preserve(box.text_frame, message)


def _replace_chart_with_insert_notice(slide, shape, chart_name: str) -> None:
    _replace_shape_with_notice(slide, shape, chart_insert_message(chart_name))


def _replace_shape_with_missing(slide, shape, what: str, *, source: str = ACTUALS_SOURCE) -> None:
    """Replace a shape with a data-not-found notice after exhaustive lookup."""
    _replace_shape_with_notice(slide, shape, data_not_found(what))


def _mark_table_banner(table, what: str, *, source: str = ACTUALS_SOURCE) -> None:
    """Put a data-not-found banner in the first cell and clear stale body values."""
    set_cell_text_preserve(table.cell(0, 0), data_not_found(what))
    if len(table.rows) > 1 or len(table.columns) > 1:
        clear_table_data_rows(table, header_rows=1)


def _clear_table_numeric_body(table, *, header_rows: int = 1, label_cols: int = 1) -> None:
    """Clear numeric columns so template values cannot persist."""
    for row_idx in range(header_rows, len(table.rows)):
        for col_idx in range(label_cols, len(table.columns)):
            set_cell_text_preserve(table.cell(row_idx, col_idx), "")


def _ensure_manual_narrative_boxes(slide) -> None:
    """Leave Leading Issues / Action Plan areas as blank textboxes (headers only)."""
    clear_manual_narrative_boxes(slide)


def _mark_gir_manual_sections(slide) -> None:
    clear_manual_narrative_boxes(slide)


def update_gir_tables(slide, data: MprData, config: dict, workbook: str) -> None:
    _strip_template_data_images(slide, label="GIR")
    fill_gir_tables(slide, data, config, workbook)
    _mark_gir_manual_sections(slide)


def update_gir_charts(slide, data: MprData, config: dict, workbook: str) -> None:
    dataset = load_gir_dataset(data, config, workbook)
    goal = dataset.mtd.goal

    _strip_template_data_images(slide, label="GIR")

    for shape in list(_iter_all_shapes(slide.shapes)):
        if shape.has_chart:
            if any(_safe_number(v) is not None for v in dataset.monthly):
                _update_chart_series(
                    slide,
                    shape,
                    dataset.monthly,
                    goal,
                    ytd_actual=dataset.ytd_actual,
                    prior_year_actuals=dataset.prior_year,
                )
            else:
                _replace_chart_with_insert_notice(slide, shape, "GIR")


def _table_has_body_data(table, *, header_rows: int = 1) -> bool:
    for row_idx in range(header_rows, len(table.rows)):
        for col_idx in range(len(table.columns)):
            text = table.cell(row_idx, col_idx).text.strip()
            if text and not is_placeholder_text(text):
                return True
    return False


def _mark_unfilled_content(slide, slide_name: str) -> None:
    """Clear non-title content on manual/placeholder slides without breaking layout."""
    for shape in list(_iter_all_shapes(slide.shapes)):
        if _is_title_shape(slide, shape):
            continue
        if shape.has_table and not _table_has_body_data(shape.table):
            clear_table_data_rows(shape.table, header_rows=1)
        elif shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text and not is_placeholder_text(text):
                clear_text_frame_content(shape.text_frame)


def _replace_largest_content_picture(slide, what: str, *, source: str = SCORECARDS_SOURCE) -> bool:
    """Replace the largest body picture with a missing-data notice."""
    pictures = [
        shape
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        and not _is_title_shape(slide, shape)
        and shape.width >= Inches(1.5)
    ]
    if not pictures:
        return False
    largest = max(pictures, key=lambda shape: shape.width * shape.height)
    _replace_shape_with_missing(slide, largest, what, source=source)
    return True


def _iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_all_shapes(shape.shapes)


def _remove_shape(shape) -> None:
    element = shape._element
    element.getparent().remove(element)


def _is_title_shape(slide, shape) -> bool:
    if slide.shapes.title is not None and shape.shape_id == slide.shapes.title.shape_id:
        return True
    if not shape.has_text_frame:
        return False
    text = shape.text_frame.text.strip().lower()
    if not text:
        return False
    if len(text) < 80 and shape.top < Pt(120):
        return True
    return False


def _agenda_line_keep(line: str, section_titles: set[str]) -> bool:
    """Return True for agenda labels that must stay (titles, times, headers)."""
    text = line.strip()
    if not text:
        return True
    lower = text.lower()
    if lower == "planned discussion":
        return True
    if lower == "gse monthly performance summary":
        return True
    if "gse monthly performance summary" in lower and len(text) < 60:
        return True
    if lower in section_titles:
        return True
    if re.match(r"^\d+\s*min$", text, re.IGNORECASE):
        return True
    if re.match(r"^\d{1,2}$", text):
        return True
    if lower in {"key", "fyi / other", "estimated discussion time", "#", "x"}:
        return True
    if lower.startswith("planned discussion") and len(text) < 40:
        return True
    return False


def _normalize_agenda_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().lower())


def _agenda_header_present(slide, header_title: str) -> bool:
    target = _normalize_agenda_text(header_title)
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        if target in _normalize_agenda_text(shape.text_frame.text):
            return True
    return False


def _ensure_agenda_header(slide, header: dict) -> None:
    """Restore item 1 title ('GSE Monthly Performance Summary') when the template slot is blank."""
    title = header["title"].strip()
    time_label = header.get("time", "").strip()
    if not title:
        return
    if _agenda_header_present(slide, title):
        return

    time_shape = None
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        if shape.text_frame.text.strip().lower() == time_label.lower():
            if time_shape is None or shape.top < time_shape.top:
                time_shape = shape

    if time_shape is not None:
        title_shape = None
        for shape in _iter_all_shapes(slide.shapes):
            if not shape.has_text_frame or shape.shape_id == time_shape.shape_id:
                continue
            if abs(shape.top - time_shape.top) > Pt(14):
                continue
            if shape.left >= time_shape.left:
                continue
            text = shape.text_frame.text.strip()
            if re.fullmatch(r"\d+", text):
                continue
            if title_shape is None or shape.left > title_shape.left:
                title_shape = shape
        if title_shape is not None:
            set_text_frame_preserve(title_shape.text_frame, title)
            logger.info("Restored agenda header title %r", title)
            return

    empty_shapes = [
        shape
        for shape in _iter_all_shapes(slide.shapes)
        if shape.has_text_frame and not shape.text_frame.text.strip()
    ]
    empty_shapes.sort(key=lambda s: (s.top, s.left))
    if empty_shapes:
        set_text_frame_preserve(empty_shapes[0].text_frame, title)
        logger.info("Restored agenda header title %r into first empty slot", title)


def update_agenda_slide(slide, element: dict) -> None:
    """Planned Discussion slide: keep section titles + times; clear bullet detail lines."""
    sections = element.get("sections") or DEFAULT_AGENDA_SECTIONS
    header = element.get("header") or DEFAULT_AGENDA_HEADER
    section_titles = {s["title"].strip().lower() for s in sections}
    section_titles.add(header["title"].strip().lower())

    cleared = kept = 0
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame
        full_text = tf.text.strip()

        # Single-line shape: keep labels, clear detail text
        if len(tf.paragraphs) == 1 or "\n" not in full_text:
            line = full_text
            if not line:
                continue
            if _agenda_line_keep(line, section_titles):
                kept += 1
                continue
            set_text_frame_preserve(tf, "")
            cleared += 1
            continue

        # Multi-line chevron shapes: keep title/time lines, clear bullet lines
        for paragraph in tf.paragraphs:
            line = paragraph.text.strip()
            if not line:
                continue
            if _agenda_line_keep(line, section_titles):
                kept += 1
            else:
                set_paragraph_text_preserve(paragraph, "")
                cleared += 1

    # Always restore item 1 header; fill other missing labels when a prior run wiped them
    _ensure_agenda_header(slide, header)
    _restore_agenda_labels(slide, sections, header, section_titles)
    logger.info("Agenda slide: kept %s label(s), cleared %s detail line(s)", kept, cleared)


def _restore_agenda_labels(
    slide,
    sections: list[dict],
    header: dict,
    section_titles: set[str],
) -> None:
    """Fill empty section label slots when a prior run wiped the agenda."""
    present_titles = set()
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            line = paragraph.text.strip()
            if line and line.lower() in section_titles:
                present_titles.add(line.lower())

    section_only = section_titles - {header["title"].strip().lower()}
    if len(present_titles & section_only) >= 3:
        return

    needed_titles = [header["title"]] + [s["title"] for s in sections]
    needed_times = [header.get("time", "")] + [s.get("time", "") for s in sections]
    needed_times = [t for t in needed_times if t]

    empty_shapes = [
        shape
        for shape in _iter_all_shapes(slide.shapes)
        if shape.has_text_frame and not shape.text_frame.text.strip()
    ]
    empty_shapes.sort(key=lambda s: (s.top, s.left))

    idx = 0
    for title in needed_titles:
        if idx >= len(empty_shapes):
            break
        set_text_frame_preserve(empty_shapes[idx].text_frame, title)
        idx += 1
    for time_label in needed_times:
        if idx >= len(empty_shapes):
            break
        set_text_frame_preserve(empty_shapes[idx].text_frame, time_label)
        idx += 1


def _replace_text(text: str, data: MprData) -> str:
    month_label = data.report_month_label()
    month_short = data.report_month_short()
    updated = text
    updated = re.sub(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b",
        month_label,
        updated,
        flags=re.IGNORECASE,
    )
    updated = re.sub(r"\b[A-Za-z]{3}'\d{2}\b", month_short, updated)
    return updated


def replace_month_tokens(prs: Presentation, data: MprData) -> None:
    for slide in prs.slides:
        for shape in _iter_all_shapes(slide.shapes):
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                if paragraph.text.strip():
                    set_paragraph_text_preserve(paragraph, _replace_text(paragraph.text, data))


def _build_chart_data(
    actuals: list[float | None],
    goal: float | None,
    existing_chart,
    *,
    ytd_actual: float | None = None,
    prior_year_actuals: list[float | None] | None = None,
) -> CategoryChartData:
    if _safe_number(ytd_actual) is not None:
        ytd = ytd_actual
    else:
        ytd_vals = [v for v in actuals[:12] if _safe_number(v) is not None]
        ytd = sum(ytd_vals) / len(ytd_vals) if ytd_vals else None
    full_actuals = list(actuals[:12]) + [ytd]
    prior_year = list((prior_year_actuals or [None] * 12)[:12]) + [None]

    chart_data = CategoryChartData()
    chart_data.categories = CHART_CATEGORIES

    if existing_chart and existing_chart.series:
        for series in existing_chart.series:
            name = series.name or ""
            lower = name.lower()
            if lower in ("actual", "current year") or "current year" in lower or lower == "actual":
                chart_data.add_series(name, tuple(full_actuals))
            elif "goal" in lower:
                chart_data.add_series(name, tuple([goal] * 13 if _safe_number(goal) is not None else [None] * 13))
            elif "p1y" in lower or "prior" in lower or "previous" in lower or "py" == lower.strip():
                chart_data.add_series(name, tuple(prior_year))
            else:
                chart_data.add_series(name, tuple([None] * 13))
    else:
        chart_data.add_series("Current Year", tuple(full_actuals))
        if _safe_number(goal) is not None:
            chart_data.add_series("Goal", tuple([goal] * 13))
        if any(_safe_number(v) is not None for v in prior_year):
            chart_data.add_series("P1Y", tuple(prior_year))

    return chart_data


def _update_chart_series(
    slide,
    shape,
    actuals: list[float | None],
    goal: float | None,
    *,
    ytd_actual: float | None = None,
    prior_year_actuals: list[float | None] | None = None,
) -> bool:
    if not shape.has_chart:
        return False
    if not any(_safe_number(v) is not None for v in actuals):
        return False

    chart_data = _build_chart_data(
        actuals,
        goal,
        shape.chart,
        ytd_actual=ytd_actual,
        prior_year_actuals=prior_year_actuals,
    )
    chart_title = ""
    chart_type = XL_CHART_TYPE.LINE_MARKERS
    if shape.has_chart:
        try:
            chart_type = shape.chart.chart_type
            if shape.chart.has_title:
                chart_title = shape.chart.chart_title.text_frame.text
        except Exception:
            pass
    try:
        shape.chart.replace_data(chart_data)
        return True
    except (ValueError, AttributeError) as exc:
        logger.info("Recreating linked/external chart: %s", exc)
        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        _remove_shape(shape)
        new_shape = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)
        if chart_title and new_shape.chart.has_title:
            set_text_frame_preserve(new_shape.chart.chart_title.text_frame, chart_title)
        return True


def _clear_non_title_content(slide, *, keep_title: bool = True, slide_name: str = "Slide") -> None:
    for shape in list(slide.shapes):
        if keep_title and _is_title_shape(slide, shape):
            continue
        if shape.has_table:
            clear_table_data_rows(shape.table, header_rows=1)
        elif shape.has_text_frame:
            clear_text_frame_content(shape.text_frame)


def _fill_table_from_dataframe(table, df, *, header_rows: int = 1) -> bool:
    if df.empty:
        return False
    max_rows = len(table.rows) - header_rows
    max_cols = len(table.columns)
    wrote = False
    rows_to_fill = min(max_rows, len(df.index))
    cols_to_fill = min(max_cols, len(df.columns))
    for r in range(rows_to_fill):
        for c in range(cols_to_fill):
            set_cell_text_preserve(table.cell(header_rows + r, c), "")
    for r in range(rows_to_fill):
        for c in range(cols_to_fill):
            val = df.iat[r, c]
            if val is None or (isinstance(val, float) and math.isnan(val)):
                continue
            text = str(val).strip()
            if text:
                set_cell_text_preserve(table.cell(header_rows + r, c), text)
                wrote = True
    return wrote


def _table_row_label(table, row_idx: int) -> str:
    return " ".join(table.cell(row_idx, col).text.strip().lower() for col in range(len(table.columns)))


def _find_table_row(table, *patterns: str) -> int | None:
    for row_idx in range(len(table.rows)):
        label = _table_row_label(table, row_idx)
        if all(pattern.lower() in label for pattern in patterns):
            return row_idx
    return None


def update_ea_asap_tables(slide, data: MprData, config: dict, workbook: str) -> None:
    """Fill EA Compliance / ASAP Reporting tables from actuals (PPT slide 6)."""
    eac_patterns = _kpi_patterns_from_config(config, "eac", "EA Compliance", "EA COMPLIANCE")
    asap_patterns = _kpi_patterns_from_config(config, "asap", "ASAP Reporting", "ASAP REPORTING")

    filled_total = fill_compliance_slide(
        slide,
        data,
        workbook=workbook,
        eac_patterns=eac_patterns,
        asap_patterns=asap_patterns,
    )

    if filled_total:
        logger.info("EA/ASAP compliance tables: wrote %s cell(s)", filled_total)
    else:
        logger.warning("EA/ASAP compliance tables: no cells filled — check KPI names in actuals")

    _ensure_manual_narrative_boxes(slide)


def update_people_table(slide, data: MprData, rows_cfg: list[dict], workbook: str) -> None:
    _strip_template_data_images(slide, label="People")
    filled_total = 0
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_table:
            continue
        filled = fill_people_table(shape.table, data, workbook=workbook)
        filled_total += filled
    if filled_total:
        logger.info("People table: wrote %s cell(s)", filled_total)
    else:
        logger.warning("People table: no cells filled — check KPI names in actuals")
    _ensure_manual_narrative_boxes(slide)


def _match_people_chart(chart_def: dict, title: str) -> bool:
    chart_title = chart_def.get("title", "")
    if chart_title and chart_title.lower() in title.lower():
        return True
    for pattern in chart_def.get("patterns", []):
        if pattern.lower() in title.lower():
            return True
    return not title.strip()


def update_people_charts(slide, data: MprData, charts_cfg: list[dict], workbook: str) -> None:
    _strip_template_data_images(slide, label="People")
    charts = [shape for shape in _iter_all_shapes(slide.shapes) if shape.has_chart]
    charts.sort(key=lambda shape: (shape.top, shape.left))
    unmatched = list(charts_cfg)

    for shape in charts:
        title = ""
        if shape.chart.has_title:
            title = shape.chart.chart_title.text_frame.text
        chart_def = None
        for candidate in unmatched:
            if _match_people_chart(candidate, title):
                chart_def = candidate
                break
        if chart_def is None and unmatched:
            chart_def = unmatched[0]

        if chart_def is None:
            _replace_chart_with_insert_notice(slide, shape, title or "People metric")
            continue

        patterns = chart_def["patterns"]
        chart_label = chart_def.get("title") or patterns[0]
        monthly, _source = lookup_monthly_series_exhaustive(data, patterns)
        mtd = lookup_kpi_exhaustive(data, patterns, month=data.month)
        goal = mtd.value.goal
        ytd_actual, _ = lookup_ytd_exhaustive(data, patterns)
        if any(_safe_number(v) is not None for v in monthly):
            _update_chart_series(
                slide,
                shape,
                monthly,
                goal,
                ytd_actual=ytd_actual,
            )
            logger.info("Updated people chart %r", chart_label)
        else:
            _replace_chart_with_insert_notice(slide, shape, chart_label)

        if chart_def in unmatched:
            unmatched.remove(chart_def)


def update_chart_slide_mapped(slide, data: MprData, charts_cfg: list[dict], workbook: str) -> None:
    _strip_template_data_images(slide, label="charts")
    charts = [s for s in list(slide.shapes) if s.has_chart]
    for shape in charts:
        title = shape.chart.chart_title.text_frame.text if shape.chart.has_title else shape.name
        matched = False
        for chart_def in charts_cfg:
            chart_title = chart_def.get("title", "")
            if chart_title and chart_title.lower() not in title.lower():
                continue
            if not chart_title and len(charts_cfg) == 1:
                pass
            elif not chart_title:
                continue
            patterns = chart_def["patterns"]
            chart_label = chart_title or patterns[0]
            monthly, _source = lookup_monthly_series_exhaustive(data, patterns)
            mtd = lookup_kpi_exhaustive(data, patterns, month=data.month)
            goal = mtd.value.goal
            if any(_safe_number(v) is not None for v in monthly):
                _update_chart_series(slide, shape, monthly, goal)
                logger.info("Updated chart %r", title)
            else:
                _replace_chart_with_insert_notice(slide, shape, chart_label)
            matched = True
            break
        if not matched and charts_cfg:
            _replace_chart_with_insert_notice(slide, shape, title or "Chart")


def _find_system_scorecard_table_shape(slide):
    """Return the shape holding the largest table on the slide."""
    candidates = [shape for shape in _iter_all_shapes(slide.shapes) if shape.has_table]
    if not candidates:
        return None
    return max(candidates, key=lambda shape: len(shape.table.rows) * len(shape.table.columns))


def _find_system_scorecard_table(slide):
    """Return the largest table on the slide (the system scorecard grid)."""
    shape = _find_system_scorecard_table_shape(slide)
    return shape.table if shape is not None else None


def _remove_content_pictures(slide) -> int:
    """Remove screenshot pictures and embedded OLE data objects from the slide body."""
    return remove_data_placeholders(slide)


def _strip_template_data_images(slide, *, label: str = "chart/table") -> int:
    """Remove large template images so stale pictures are not shown as data."""
    removed = _remove_content_pictures(slide)
    if removed:
        logger.info("Removed %s template image(s) from slide body (%s)", removed, label)
    return removed


def _fill_scorecard_table_from_block(table, block, *, fill_mode: str = "full") -> int:
    """Copy scorecard data into the template table, preserving cell fonts and colors."""
    if block.empty:
        return 0

    excel_header_row = find_scorecard_header_row(block)
    table_header_row = find_scorecard_table_header_row(table)
    table_headers = [table.cell(table_header_row, col).text.strip() for col in range(len(table.columns))]

    col_map = scorecard_column_map(block, excel_header_row, table_headers) if excel_header_row is not None else {}
    if not col_map:
        col_map = {col: col for col in range(min(len(block.columns), len(table.columns)))}

    filled = 0

    if fill_mode == "full":
        table_start = 0
        for excel_row in range(len(block)):
            ppt_row = table_start + excel_row
            if ppt_row >= len(table.rows):
                break
            for excel_col, table_col in col_map.items():
                if table_col >= len(table.columns):
                    continue
                value = format_scorecard_cell(block.iat[excel_row, excel_col])
                if value:
                    set_cell_text_preserve(table.cell(ppt_row, table_col), value)
                    filled += 1
        return filled

    excel_index = build_scorecard_row_index(block, header_row=excel_header_row)
    body_start = table_header_row + 1
    section = ""
    kpi = ""

    for ppt_row in range(body_start, len(table.rows)):
        ppt_values = [table.cell(ppt_row, col).text.strip() for col in range(len(table.columns))]
        if not any(ppt_values):
            continue

        section, kpi, key = contextual_scorecard_key(ppt_values, current_section=section, current_kpi=kpi)
        if not key:
            continue

        excel_row = excel_index.get(key)
        if excel_row is None:
            continue

        for excel_col, table_col in col_map.items():
            if table_col >= len(table.columns):
                continue
            value = format_scorecard_cell(block.iat[excel_row, excel_col])
            if value:
                set_cell_text_preserve(table.cell(ppt_row, table_col), value)
                filled += 1

    if filled == 0:
        data_start = (excel_header_row + 1) if excel_header_row is not None else 0
        for offset, excel_row in enumerate(range(data_start, len(block))):
            ppt_row = body_start + offset
            if ppt_row >= len(table.rows):
                break
            for excel_col, table_col in col_map.items():
                if table_col >= len(table.columns):
                    continue
                value = format_scorecard_cell(block.iat[excel_row, excel_col])
                if value:
                    set_cell_text_preserve(table.cell(ppt_row, table_col), value)
                    filled += 1

    return filled


def _trim_sheet_for_paste(df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """Drop empty trailing rows/columns before pasting; return trim offsets."""
    if df.empty:
        return df, 0, 0
    row_offset = 0
    col_offset = 0
    out = df.copy()
    while len(out.columns) > 0 and out.iloc[:, -1].isna().all():
        out = out.iloc[:, :-1]
    while len(out.columns) > 0 and out.iloc[:, 0].isna().all():
        out = out.iloc[:, 1:]
        col_offset += 1
    while len(out.index) > 0 and out.iloc[-1].isna().all():
        out = out.iloc[:-1]
    while len(out.index) > 0 and out.iloc[0].isna().all():
        out = out.iloc[1:]
        row_offset += 1
    return out.reset_index(drop=True), row_offset, col_offset


def update_system_scorecard_table(slide, data: MprData, element: dict) -> bool:
    """Fill a System Scorecard slide from the scorecards workbook (paste native table)."""
    workbook = element.get("workbook", "scorecards")
    block_name = element.get("block", "summary_1")
    block = load_system_scorecard(
        data,
        workbook=workbook,
        sheet_name=element.get("sheet"),
        block=block_name,
    )

    if block.empty:
        logger.warning("System scorecard %s: no data extracted from %s — leaving template", block_name, workbook)
        return False

    table_shape = _find_system_scorecard_table_shape(slide)
    table = table_shape.table if table_shape is not None else None
    if table is None:
        _, filled = paste_dataframe_on_slide(slide, block, style="scorecard")
        if filled:
            logger.info(
                "System scorecard %s: pasted %s cells into new native table from %s rows",
                block_name,
                filled,
                len(block),
            )
            return True
        logger.warning("System scorecard %s: no placeholder found to replace with table", block_name)
        return False

    removed = _remove_content_pictures(slide)
    if removed:
        logger.info("Removed %s scorecard image(s); using native template table", removed)

    fill_mode = element.get("fill_mode", "full")
    filled = _fill_scorecard_table_from_block(table, block, fill_mode=fill_mode)
    if filled == 0:
        logger.warning("System scorecard %s: no cells written — leaving template table", block_name)
        return False
    fit_scorecard_table(
        table,
        block,
        width_emu=table_shape.width,
        height_emu=table_shape.height,
    )
    logger.info(
        "System scorecard %s table: wrote %s cell(s) from %s rows (mode=%s)",
        block_name,
        filled,
        len(block),
        fill_mode,
    )
    return filled > 0


def apply_scorecard_sheet(slide, data: MprData, element: dict) -> bool:
    """Paste an entity scorecard sheet as a native table (replaces screenshot)."""
    workbook = element["workbook"]
    sheet_index = element.get("sheet_index", 0)
    sheet_name = element.get("sheet")
    if sheet_name:
        raw = data.store.read_sheet(workbook, sheet_name, raw=True)
    else:
        names = data.sheet_names(workbook)
        if sheet_index >= len(names):
            return False
        raw = data.store.read_sheet(workbook, names[sheet_index], raw=True)

    df, row_offset, col_offset = _trim_sheet_for_paste(raw)
    if df.empty:
        logger.warning("Scorecard sheet empty for slide element — leaving template")
        return False

    ws = None
    try:
        ws = data.store.open_worksheet(workbook, sheet_name or data.sheet_names(workbook)[sheet_index])
    except (KeyError, FileNotFoundError, ValueError):
        ws = None

    for shape in slide.shapes:
        if shape.has_table:
            if ws is not None:
                filled = fill_styled_worksheet_table(
                    shape.table, df, ws, origin_row=row_offset, origin_col=col_offset
                )
            else:
                filled = fill_styled_scorecard_table(shape.table, df)
            if filled:
                _remove_content_pictures(slide)
                clear_manual_narrative_boxes(slide)
                return True

    if ws is not None:
        table, filled = paste_dataframe_on_slide(
            slide,
            df,
            style="worksheet",
            ws=ws,
            origin_row=row_offset,
            origin_col=col_offset,
        )
    else:
        table, filled = paste_dataframe_on_slide(slide, df, style="scorecard")
    if filled:
        logger.info("Entity scorecard: pasted %s cells (%s rows) into styled native table", filled, len(df))
        clear_manual_narrative_boxes(slide)
        return True
    logger.warning("Scorecard sheet: no screenshot placeholder to replace — leaving template")
    return False


def apply_workings_table(slide, data: MprData, element: dict) -> bool:
    """Paste finance/workings data as a native table (replaces screenshot on Finance slide)."""
    workbook = element["workbook"]
    sheet = element.get("sheet")
    if sheet:
        raw = data.store.read_sheet(workbook, sheet, raw=True)
    else:
        names = data.sheet_names(workbook)
        if not names:
            return False
        raw = data.store.read_sheet(workbook, names[0], raw=True)

    df, row_offset, col_offset = _trim_sheet_for_paste(raw)
    if df.empty:
        logger.warning("Workings sheet empty — leaving template")
        return False

    sheet_name = sheet or data.sheet_names(workbook)[0]
    ws = None
    try:
        ws = data.store.open_worksheet(workbook, sheet_name)
    except (KeyError, FileNotFoundError, ValueError):
        ws = None

    for shape in slide.shapes:
        if shape.has_table:
            if ws is not None:
                filled = fill_styled_worksheet_table(
                    shape.table, df, ws, origin_row=row_offset, origin_col=col_offset
                )
            else:
                filled = fill_styled_scorecard_table(shape.table, df)
            if filled:
                _remove_content_pictures(slide)
                clear_manual_narrative_boxes(slide)
                return True

    if ws is not None:
        table, filled = paste_dataframe_on_slide(
            slide,
            df,
            style="worksheet",
            ws=ws,
            origin_row=row_offset,
            origin_col=col_offset,
        )
    else:
        table, filled = paste_dataframe_on_slide(slide, df, style="scorecard")
    if filled:
        logger.info("Finance/workings: pasted %s cells into styled native table", filled)
        clear_manual_narrative_boxes(slide)
        return True
    logger.warning("Workings sheet: no screenshot placeholder to replace — leaving template")
    return False


def _apply_element(slide, data: MprData, config: dict, element: dict) -> None:
    etype = element["type"]
    workbook = element.get("workbook", "actuals")

    if etype == "month_tokens":
        return
    if etype == "gir_tables":
        update_gir_tables(slide, data, config, workbook)
    elif etype == "gir_charts":
        update_gir_charts(slide, data, config, workbook)
    elif etype == "ea_asap_tables":
        update_ea_asap_tables(slide, data, config, workbook)
    elif etype == "people_table":
        update_people_table(slide, data, element.get("rows", []), workbook)
    elif etype == "people_charts":
        update_people_charts(slide, data, element.get("charts", []), workbook)
    elif etype == "chart_slide":
        update_chart_slide_mapped(slide, data, element.get("charts", []), workbook)
    elif etype == "system_scorecard_table":
        if not update_system_scorecard_table(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No system scorecard table data for slide element %s", element)
    elif etype == "scorecard_sheet":
        if not apply_scorecard_sheet(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No scorecard data for slide element %s", element)
    elif etype == "workings_table":
        if not apply_workings_table(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No workings data for slide element %s", element)
    elif etype == "agenda_slide":
        update_agenda_slide(slide, element)


def _apply_slide_spec(slide, data: MprData, config: dict, slide_spec: dict) -> None:
    if slide_spec.get("clear_pictures"):
        for shape in list(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and not _is_title_shape(slide, shape):
                _remove_shape(shape)

    for element in slide_spec.get("elements", []):
        _apply_element(slide, data, config, element)

    if slide_spec.get("clear_non_titles"):
        _mark_unfilled_content(slide, slide_spec.get("name", "Slide"))

    clear_manual_narrative_boxes(slide)


def load_template_map(base_dir: Path) -> list[dict]:
    path = base_dir / "template_map.yaml"
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as handle:
        payload = yaml.safe_load(handle) or {}
    return payload.get("slides", [])


def _load_template(config: dict, base_dir: Path) -> Presentation:
    template_rel = config["powerpoint"]["template_path"]
    cached = config.get("_sharepoint_files", {}).get(template_rel)
    if cached:
        logger.info("Using PowerPoint template from SharePoint memory (%s bytes)", len(cached))
        return Presentation(io.BytesIO(cached))

    from sharepoint_live import get_cached_file

    cached = get_cached_file(template_rel)
    if cached:
        return Presentation(io.BytesIO(cached))

    template_path = base_dir / template_rel
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    return Presentation(str(template_path))


def _save_presentation(prs: Presentation, output_path: Path) -> Path:
    """Save PPT to a fixed output path, replacing any existing file with the same name."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        try:
            output_path.unlink()
            logger.info("Removed previous report: %s", output_path)
        except PermissionError as exc:
            raise PermissionError(
                f"Cannot replace '{output_path.name}' — close it in PowerPoint or Edge and re-run."
            ) from exc

    try:
        prs.save(str(output_path))
    except PermissionError as exc:
        raise PermissionError(
            f"Cannot save '{output_path.name}' — close it in PowerPoint or Edge and re-run."
        ) from exc

    return output_path


def build_presentation(data: MprData, config: dict, base_dir: Path) -> Path:
    ppt_cfg = config["powerpoint"]
    month_name = date(data.year, data.month, 1).strftime("%B")
    if ppt_cfg.get("output_name"):
        output_name = ppt_cfg["output_name"].format(
            year=data.year,
            month=data.month,
            month_name=month_name,
            report_title=data.report_output_title(),
        )
    else:
        output_name = f"{data.report_output_title()}.pptx"
    output_dir = base_dir / ppt_cfg["output_dir"]
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / output_name

    prs = _load_template(config, base_dir)
    replace_month_tokens(prs, data)

    slide_specs = load_template_map(base_dir)
    if slide_specs:
        spec_by_index = {spec["index"]: spec for spec in slide_specs}
        for idx, slide in enumerate(prs.slides):
            spec = spec_by_index.get(idx)
            if spec:
                _apply_slide_spec(slide, data, config, spec)
                logger.info("Processed slide %s (%s)", idx, spec.get("name", ""))
            elif config.get("powerpoint", {}).get("clear_unmapped_slides", False):
                _clear_non_title_content(slide, slide_name=f"Slide {idx + 1}")
    else:
        logger.warning("template_map.yaml not found — only month tokens were updated.")

    for slide in prs.slides:
        clear_manual_narrative_boxes(slide)

    saved_path = _save_presentation(prs, output_path)
    logger.info("Saved report: %s", saved_path)
    return saved_path
