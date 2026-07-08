"""Fill the GSE MPR PowerPoint template with Excel data."""

from __future__ import annotations

import logging
import re
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE

from mpr_data import MONTH_LABELS, MprData

logger = logging.getLogger(__name__)
CHART_CATEGORIES = MONTH_LABELS + ["YTD"]


def _fmt(value, decimals=2):
    if value is None:
        return ""
    return f"{round(value):.0f}" if decimals == 0 else f"{value:.{decimals}f}"


def _fmt_diff(actual, goal, decimals=2):
    if actual is None or goal is None:
        return ""
    diff = actual - goal
    return f"({abs(diff):.{decimals}f})" if diff < 0 else _fmt(diff, decimals)


def _set_cell(table, row, col, value):
    if row < len(table.rows) and col < len(table.columns):
        table.cell(row, col).text = value


def _replace_text(text, data):
    month_label = data.report_month_label()
    month_short = data.report_month_short()
    updated = re.sub(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\b",
        month_label, text, flags=re.IGNORECASE,
    )
    return updated.replace("May'26", month_short).replace("May 2026", month_label)


def _iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_all_shapes(shape.shapes)


def replace_month_tokens(prs, data):
    for slide in prs.slides:
        for shape in _iter_all_shapes(slide.shapes):
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    if p.text.strip():
                        p.text = _replace_text(p.text, data)


def _update_chart_series(chart, actuals, goal):
    try:
        ytd_vals = [v for v in actuals[:12] if v is not None]
        ytd = sum(ytd_vals) / len(ytd_vals) if ytd_vals else None
        full_actuals = list(actuals[:12]) + [ytd]
        chart_data = CategoryChartData()
        chart_data.categories = CHART_CATEGORIES
        for series in chart.series:
            name = (series.name or "").lower()
            if "actual" in name or name == "current year":
                chart_data.add_series(series.name, tuple(full_actuals))
            elif "goal" in name:
                chart_data.add_series(series.name, tuple([goal] * 13 if goal else [None] * 13))
            else:
                existing = list(series.values)
                while len(existing) < 13:
                    existing.append(None)
                chart_data.add_series(series.name, tuple(existing[:13]))
        if not chart.series:
            chart_data.add_series("Actual", tuple(full_actuals))
            if goal:
                chart_data.add_series("Goal", tuple([goal] * 13))
        chart.replace_data(chart_data)
    except ValueError as exc:
        logger.warning("Skipped chart update: %s", exc)


def update_gir_slide(slide, data):
    mtd = data.kpi_value(["GIR"], month=data.month)
    monthly = data.monthly_series(["GIR"], through_month=data.month)
    ytd_actual = data.ytd_value(["GIR"])
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        label = table.cell(0, 0).text.strip().lower()
        if label == "metric" and table.cell(2, 0).text.strip().upper() == "GIR":
            _set_cell(table, 2, 1, _fmt(mtd.actual))
            _set_cell(table, 2, 2, _fmt(mtd.goal))
            _set_cell(table, 2, 3, _fmt_diff(mtd.actual, mtd.goal))
            _set_cell(table, 2, 4, _fmt(ytd_actual))
            _set_cell(table, 2, 5, _fmt(mtd.goal))
            _set_cell(table, 2, 6, _fmt_diff(ytd_actual, mtd.goal))
        if "actual:" in table.cell(1, 0).text.lower():
            _set_cell(table, 0, 0, data.report_month_short())
            _set_cell(table, 1, 1, _fmt(mtd.actual))
            if mtd.goal:
                _set_cell(table, 2, 1, _fmt(mtd.goal))
        if label == "recordable":
            for col_idx, month_num in enumerate(range(1, 13), start=1):
                if month_num <= data.month:
                    _set_cell(table, 1, col_idx, _fmt(monthly[month_num - 1]))
            _set_cell(table, 1, 13, _fmt(ytd_actual))
    for shape in slide.shapes:
        if shape.has_chart:
            _update_chart_series(shape.chart, monthly, mtd.goal)


def update_chart_slide(slide, data, mappings):
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        title = shape.chart.chart_title.text_frame.text if shape.chart.has_title else shape.name
        for chart_title, patterns in mappings:
            if chart_title.lower() in title.lower() or not chart_title:
                monthly = data.monthly_series(patterns, through_month=data.month)
                goal = data.kpi_value(patterns, month=data.month).goal
                _update_chart_series(shape.chart, monthly, goal)
                break


def build_presentation(data, config, base_dir):
    ppt_cfg = config["powerpoint"]
    template_path = base_dir / ppt_cfg["template_path"]
    output_path = base_dir / ppt_cfg["output_dir"] / ppt_cfg["output_name"].format(year=data.year, month=data.month)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(template_path))
    replace_month_tokens(prs, data)
    slide_map = config.get("slides", {})
    updates = [
        (slide_map.get("gir", 4), lambda s: update_gir_slide(s, data), "GIR"),
        (slide_map.get("pmi", 10), lambda s: update_chart_slide(s, data, [
            ("Motorized", ["PM (M)"]),
            ("Stationary", ["PM (S)"]),
        ]), "PMI"),
        (slide_map.get("isr", 11), lambda s: update_chart_slide(s, data, [
            ("Reliability", ["REL"]),
            ("Severity", ["SEV"]),
        ]), "ISR"),
        (slide_map.get("operations", 19), lambda s: update_chart_slide(s, data, [
            ("Jam Rate", ["Jams"]),
            ("Clear Times", ["Times"]),
        ]), "Ops"),
        (slide_map.get("vos", 21), lambda s: update_chart_slide(s, data, [
            ("", ["VOS (S)"]),
        ]), "VOS"),
    ]
    for idx, updater, label in updates:
        if idx < len(prs.slides):
            try:
                updater(prs.slides[idx])
                logger.info("Updated %s slide (%s)", label, idx)
            except Exception as exc:
                logger.warning("Could not update %s slide: %s", label, exc)
    prs.save(output_path)
    return output_path