"""Fill the GSE MPR PowerPoint template with Excel data."""

from __future__ import annotations

import io
import logging
import math
import re
from datetime import date
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

from mpr_data import MONTH_LABELS, MprData
from ppt_format import (
    clear_table_data_rows,
    clear_text_frame_content,
    set_cell_text_preserve,
    set_paragraph_text_preserve,
    set_text_frame_preserve,
)

logger = logging.getLogger(__name__)

CHART_CATEGORIES = MONTH_LABELS + ["YTD"]


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


def replace_month_tokens_on_slide(slide, data: MprData) -> None:
    """Update month/year tokens on a single slide only."""
    for shape in _iter_all_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            original = paragraph.text
            if not original.strip():
                continue
            updated = _replace_text(original, data)
            if updated != original:
                set_paragraph_text_preserve(paragraph, updated)


# Planned Discussion (slide 2) — static agenda matching the GSE MPR template.
AGENDA_TOPICS: list[tuple[str, str]] = [
    ("Safety", "10 min"),
    ("People", "10 min"),
    ("Finance", "10 min"),
    ("Customer Experience", "10 min"),
    ("North (M)", "10 min"),
    ("South (M)", "10 min"),
    ("Stationary", "15 min"),
    ("Galley", "10 min"),
    ("Closing", "3 min"),
]


def fill_agenda_slide(slide) -> None:
    """Restore Planned Discussion topic titles and time ovals left-to-right.

    The template uses separate text boxes (titles) and bottom ovals (times).
    Older builds cleared those; this writes the canonical agenda back in.
    """
    title_boxes = [
        shape
        for shape in slide.shapes
        if shape.has_text_frame
        and 1_700_000 <= int(shape.top) <= 2_100_000
        and int(shape.width) > 500_000
    ]
    time_ovals = [
        shape
        for shape in slide.shapes
        if shape.has_text_frame
        and "oval" in (shape.name or "").lower()
        and int(shape.top) > 4_700_000
    ]
    title_boxes.sort(key=lambda s: int(s.left))
    time_ovals.sort(key=lambda s: int(s.left))

    if len(title_boxes) < len(AGENDA_TOPICS):
        logger.warning(
            "Agenda slide: expected %s topic title boxes, found %s",
            len(AGENDA_TOPICS),
            len(title_boxes),
        )
    if len(time_ovals) < len(AGENDA_TOPICS):
        logger.warning(
            "Agenda slide: expected %s time ovals, found %s",
            len(AGENDA_TOPICS),
            len(time_ovals),
        )

    for idx, (topic, minutes) in enumerate(AGENDA_TOPICS):
        if idx < len(title_boxes):
            set_text_frame_preserve(title_boxes[idx].text_frame, topic)
        if idx < len(time_ovals):
            set_text_frame_preserve(time_ovals[idx].text_frame, minutes)

    logger.info("Filled Planned Discussion agenda (%s topics)", len(AGENDA_TOPICS))


def _build_chart_data(actuals: list[float | None], goal: float | None, existing_chart) -> CategoryChartData:
    ytd_vals = [v for v in actuals[:12] if _safe_number(v) is not None]
    ytd = sum(ytd_vals) / len(ytd_vals) if ytd_vals else None
    full_actuals = list(actuals[:12]) + [ytd]

    chart_data = CategoryChartData()
    chart_data.categories = CHART_CATEGORIES

    if existing_chart and existing_chart.series:
        for series in existing_chart.series:
            name = series.name or ""
            lower = name.lower()
            if lower in ("actual", "current year") or "actual" in lower:
                chart_data.add_series(name, tuple(full_actuals))
            elif "goal" in lower:
                chart_data.add_series(name, tuple([goal] * 13 if _safe_number(goal) is not None else [None] * 13))
            else:
                existing = list(series.values)
                while len(existing) < 13:
                    existing.append(None)
                chart_data.add_series(name, tuple(existing[:13]))
    else:
        chart_data.add_series("Actual", tuple(full_actuals))
        if _safe_number(goal) is not None:
            chart_data.add_series("Goal", tuple([goal] * 13))

    return chart_data


def _update_chart_series(slide, shape, actuals: list[float | None], goal: float | None) -> bool:
    if not shape.has_chart:
        return False
    if not any(_safe_number(v) is not None for v in actuals):
        return False

    chart_data = _build_chart_data(actuals, goal, shape.chart)
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


def _clear_non_title_content(slide, *, keep_title: bool = True) -> None:
    for shape in list(slide.shapes):
        if keep_title and _is_title_shape(slide, shape):
            continue
        if shape.has_table:
            clear_table_data_rows(shape.table, header_rows=1)
        elif shape.has_chart:
            _remove_shape(shape)
        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            _remove_shape(shape)
        elif shape.has_text_frame:
            clear_text_frame_content(shape.text_frame)


def _fill_table_from_dataframe(table, df, *, header_rows: int = 1) -> bool:
    if df.empty:
        return False
    max_rows = len(table.rows) - header_rows
    max_cols = len(table.columns)
    wrote = False
    for r in range(min(max_rows, len(df.index))):
        for c in range(min(max_cols, len(df.columns))):
            val = df.iat[r, c]
            if val is None or (isinstance(val, float) and math.isnan(val)):
                text = ""
            else:
                text = str(val).strip()
            if text:
                set_cell_text_preserve(table.cell(header_rows + r, c), text)
                wrote = True
    return wrote


def update_gir_tables(slide, data: MprData, config: dict, workbook: str) -> None:
    kpi = config.get("kpi_mappings", {})
    gir_name = kpi.get("gir", "GIR")
    injury_name = kpi.get("injury_count", "Injury Count")

    mtd = data.kpi_value([gir_name], month=data.month, workbook=workbook)
    ytd_actual = data.ytd_value([gir_name], workbook=workbook)
    monthly = data.monthly_series([gir_name], through_month=data.month, workbook=workbook)
    injury_mtd = data.kpi_value([injury_name], month=data.month, workbook=workbook)
    injury_monthly = data.monthly_series([injury_name], through_month=data.month, workbook=workbook)
    injury_ytd = data.ytd_value([injury_name], workbook=workbook)

    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        label = table.cell(0, 0).text.strip().lower()

        if label == "metric" and table.cell(2, 0).text.strip().upper() == "GIR":
            set_cell_text_preserve(table.cell(2, 1), _fmt(mtd.actual))
            set_cell_text_preserve(table.cell(2, 2), _fmt(mtd.goal))
            set_cell_text_preserve(table.cell(2, 3), _fmt_diff(mtd.actual, mtd.goal))
            set_cell_text_preserve(table.cell(2, 4), _fmt(ytd_actual))
            set_cell_text_preserve(table.cell(2, 5), _fmt(mtd.goal))
            set_cell_text_preserve(table.cell(2, 6), _fmt_diff(ytd_actual, mtd.goal))

        if "actual:" in table.cell(1, 0).text.lower() or data.report_month_short().lower() in table.cell(0, 0).text.lower():
            set_cell_text_preserve(table.cell(0, 0), data.report_month_short())
            set_cell_text_preserve(table.cell(1, 1), _fmt(mtd.actual))
            set_cell_text_preserve(table.cell(2, 1), _fmt(mtd.goal))

        if label == "recordable":
            for col_idx, month_num in enumerate(range(1, 13), start=1):
                if month_num <= data.month:
                    set_cell_text_preserve(table.cell(1, col_idx), _fmt(monthly[month_num - 1]))
            set_cell_text_preserve(table.cell(1, 13), _fmt(ytd_actual))
            set_cell_text_preserve(table.cell(1, 14), _fmt_diff(ytd_actual, mtd.goal))
            if len(table.rows) > 2 and "injury" in table.cell(2, 0).text.lower():
                for col_idx, month_num in enumerate(range(1, 13), start=1):
                    if month_num <= data.month and _safe_number(injury_monthly[month_num - 1]) is not None:
                        set_cell_text_preserve(table.cell(2, col_idx), _fmt(injury_monthly[month_num - 1], decimals=0))
                if _safe_number(injury_ytd) is not None:
                    set_cell_text_preserve(table.cell(2, 13), _fmt(injury_ytd, decimals=0))


def update_gir_charts(slide, data: MprData, config: dict, workbook: str) -> None:
    gir_name = config.get("kpi_mappings", {}).get("gir", "GIR")
    monthly = data.monthly_series([gir_name], through_month=data.month, workbook=workbook)
    goal = data.kpi_value([gir_name], month=data.month, workbook=workbook).goal
    for shape in list(slide.shapes):
        if shape.has_chart:
            if any(_safe_number(v) is not None for v in monthly):
                _update_chart_series(slide, shape, monthly, goal)
            else:
                _remove_shape(shape)


def update_people_table(slide, data: MprData, rows_cfg: list[dict], workbook: str) -> None:
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        for row_idx in range(2, len(table.rows)):
            label = table.cell(row_idx, 0).text.strip().upper()
            for row_def in rows_cfg:
                if row_def["label"] in label:
                    patterns = row_def["patterns"]
                    mtd = data.kpi_value(patterns, month=data.month, workbook=workbook)
                    ytd = data.ytd_value(patterns, workbook=workbook)
                    set_cell_text_preserve(
                        table.cell(row_idx, 1),
                        _fmt(mtd.actual, 0) if _safe_number(mtd.actual) is not None else "",
                    )
                    set_cell_text_preserve(
                        table.cell(row_idx, 3),
                        _fmt(ytd, 0) if _safe_number(ytd) is not None else "",
                    )
                    if _safe_number(mtd.goal) is not None and _safe_number(ytd) is not None:
                        set_cell_text_preserve(table.cell(row_idx, 4), _fmt_diff(ytd, mtd.goal, 0))
                    else:
                        set_cell_text_preserve(table.cell(row_idx, 4), "")
                    break


def update_people_charts(slide, data: MprData, charts_cfg: list[dict], workbook: str) -> None:
    for shape in list(slide.shapes):
        if not shape.has_chart:
            continue
        title = shape.chart.chart_title.text_frame.text if shape.chart.has_title else ""
        for chart_def in charts_cfg:
            if chart_def["title"].lower() in title.lower():
                monthly = data.monthly_series(chart_def["patterns"], through_month=data.month, workbook=workbook)
                goal = data.kpi_value(chart_def["patterns"], month=data.month, workbook=workbook).goal
                if any(_safe_number(v) is not None for v in monthly):
                    _update_chart_series(slide, shape, monthly, goal)
                else:
                    _remove_shape(shape)
                break


def update_chart_slide_mapped(slide, data: MprData, charts_cfg: list[dict], workbook: str) -> None:
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
            monthly = data.monthly_series(chart_def["patterns"], through_month=data.month, workbook=workbook)
            goal = data.kpi_value(chart_def["patterns"], month=data.month, workbook=workbook).goal
            if any(_safe_number(v) is not None for v in monthly):
                _update_chart_series(slide, shape, monthly, goal)
                logger.info("Updated chart %r", title)
            else:
                _remove_shape(shape)
            matched = True
            break
        if not matched and charts_cfg:
            _remove_shape(shape)


def apply_scorecard_sheet(slide, data: MprData, element: dict) -> bool:
    workbook = element["workbook"]
    sheet_index = element.get("sheet_index", 0)
    sheet_name = element.get("sheet")
    if sheet_name:
        df = data.read_sheet(workbook, sheet_name)
    else:
        names = data.sheet_names(workbook)
        if sheet_index >= len(names):
            return False
        df = data.read_sheet(workbook, names[sheet_index])

    if df.empty:
        return False

    for shape in slide.shapes:
        if shape.has_table:
            if _fill_table_from_dataframe(shape.table, df.head(len(shape.table.rows) - 1)):
                return True
    return False


def apply_scorecard_screenshot_element(slide, data: MprData, element: dict) -> bool:
    from scorecard_screenshots import apply_scorecard_screenshot

    return apply_scorecard_screenshot(slide, data, element)


def apply_workings_table(slide, data: MprData, element: dict) -> bool:
    workbook = element["workbook"]
    sheet = element.get("sheet")
    if sheet:
        df = data.read_sheet(workbook, sheet)
    else:
        names = data.sheet_names(workbook)
        if not names:
            return False
        df = data.read_sheet(workbook, names[0])
    if df.empty:
        return False
    for shape in slide.shapes:
        if shape.has_table:
            if _fill_table_from_dataframe(shape.table, df):
                return True
    return False


def _apply_element(slide, data: MprData, config: dict, element: dict) -> None:
    etype = element["type"]
    workbook = element.get("workbook", "actuals")

    if etype == "month_tokens":
        replace_month_tokens_on_slide(slide, data)
    elif etype == "agenda":
        fill_agenda_slide(slide)
    elif etype == "gir_tables":
        update_gir_tables(slide, data, config, workbook)
    elif etype == "gir_charts":
        update_gir_charts(slide, data, config, workbook)
    elif etype == "people_table":
        update_people_table(slide, data, element.get("rows", []), workbook)
    elif etype == "people_charts":
        update_people_charts(slide, data, element.get("charts", []), workbook)
    elif etype == "chart_slide":
        update_chart_slide_mapped(slide, data, element.get("charts", []), workbook)
    elif etype == "scorecard_sheet":
        if not apply_scorecard_sheet(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No scorecard data for slide element %s", element)
    elif etype == "scorecard_screenshot":
        if not apply_scorecard_screenshot_element(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No System scorecard screenshot for element %s", element)
    elif etype == "workings_table":
        if not apply_workings_table(slide, data, element):
            if not element.get("optional", False):
                logger.warning("No workings data for slide element %s", element)


def _apply_slide_spec(slide, data: MprData, config: dict, slide_spec: dict) -> None:
    """Apply mapped updates while preserving the rest of the template slide."""
    # Never wipe template chrome. Only touch mapped elements.
    for element in slide_spec.get("elements", []):
        _apply_element(slide, data, config, element)


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


def build_presentation(data: MprData, config: dict, base_dir: Path) -> Path:
    ppt_cfg = config["powerpoint"]
    month_name = date(data.year, data.month, 1).strftime("%B")
    output_name = ppt_cfg.get("output_name", "GSE MPR - {month_name} {year}.pptx").format(
        year=data.year,
        month=data.month,
        month_name=month_name,
        report_title=data.report_output_title(),
    )
    output_dir = base_dir / ppt_cfg["output_dir"]
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / output_name

    prs = _load_template(config, base_dir)

    slide_specs = load_template_map(base_dir)
    if slide_specs:
        spec_by_index = {spec["index"]: spec for spec in slide_specs}
        for idx, slide in enumerate(prs.slides):
            spec = spec_by_index.get(idx)
            if spec:
                _apply_slide_spec(slide, data, config, spec)
                logger.info("Processed slide %s (%s)", idx, spec.get("name", ""))
            elif ppt_cfg.get("clear_unmapped_slides", False):
                _clear_non_title_content(slide)
    else:
        logger.warning("template_map.yaml not found — applying month tokens on slide 0 only.")
        if prs.slides:
            replace_month_tokens_on_slide(prs.slides[0], data)

    prs.save(output_path)
    logger.info("Saved report: %s", output_path)
    return output_path