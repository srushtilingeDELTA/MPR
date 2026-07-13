"""Fill Leadership Engagement / Psychological Safety / Accountability slide."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from data_lookup import lookup_kpi_exhaustive, lookup_ytd_exhaustive
from ppt_format import set_cell_text_preserve
from ppt_missing import EMPTY_CELL, data_not_found, has_numeric_value
from scorecard_data import (
    build_scorecard_row_index,
    find_scorecard_header_row,
    format_scorecard_cell,
    load_system_scorecard,
)

if TYPE_CHECKING:
    from mpr_data import MprData

ACTUALS_SOURCE = "MPR Actuals and Goals_v2.xlsx"
SCORECARDS_SOURCE = "2026 - GSE Scorecards.xlsx"

PEOPLE_SECTIONS = {
    "LEADERSHIP ENGAGEMENT": ["Leadership Engagement", "LEADERSHIP ENGAGEMENT"],
    "PSYCHOLOGICAL SAFETY": ["Psychological Safety", "PSYCHOLOGICAL SAFETY"],
    "ACCOUNTABILITY": ["Accountability", "ACCOUNTABILITY"],
}

PEOPLE_CHART_ORDER = [
    "Leadership Engagement",
    "Psychological Safety",
    "Accountability",
]

METRIC_KEYS = ("mtd_actual", "mtd_bw", "ytd_actual", "ytd_bw", "score_mtd", "score_ytd")


def _normalize_label(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().upper())


def is_people_table(table) -> bool:
    text = " ".join(
        table.cell(row, col).text.strip().lower()
        for row in range(min(4, len(table.rows)))
        for col in range(min(4, len(table.columns)))
    )
    if "leadership engagement" in text or "psychological safety" in text:
        return True
    labels = {_normalize_label(table.cell(row, 0).text) for row in range(len(table.rows))}
    return bool(labels & set(PEOPLE_SECTIONS))


def _column_header_context(table, col_idx: int) -> str:
    parts = []
    for row_idx in range(min(3, len(table.rows))):
        text = table.cell(row_idx, col_idx).text.strip().lower()
        if text:
            parts.append(text)
    for header_row in (0, 1):
        if header_row >= len(table.rows):
            continue
        for scan_col in range(col_idx, -1, -1):
            group = table.cell(header_row, scan_col).text.strip().lower()
            if group in {"mtd", "ytd", "score"}:
                parts.insert(0, group)
                break
        if parts and parts[0] in {"mtd", "ytd", "score"}:
            break
    return " | ".join(parts)


def _metric_from_label(label: str) -> str | None:
    if "score" in label:
        return "score_ytd" if "ytd" in label else "score_mtd"
    if "mtd" in label:
        if "actual" in label:
            return "mtd_actual"
        if "b/(w)" in label or "b/w" in label or "goal" in label:
            return "mtd_bw"
    if "ytd" in label:
        if "actual" in label:
            return "ytd_actual"
        if "b/(w)" in label or "b/w" in label or "goal" in label:
            return "ytd_bw"
    if label.strip() == "actual":
        return "mtd_actual"
    return None


def parse_people_table_layout(table) -> dict[str, int]:
    layout: dict[str, int] = {}
    for col_idx in range(1, len(table.columns)):
        metric = _metric_from_label(_column_header_context(table, col_idx))
        if metric and metric not in layout:
            layout[metric] = col_idx
    return layout


def _kpi_patterns_for_row(label: str) -> list[str]:
    normalized = _normalize_label(label)
    if normalized in PEOPLE_SECTIONS:
        return list(PEOPLE_SECTIONS[normalized])
    text = label.strip()
    if not text:
        return []
    patterns = [text]
    if len(text) > 45:
        patterns.append(text[:45].rstrip())
    first_sentence = text.split(".")[0].strip()
    if first_sentence and first_sentence not in patterns:
        patterns.append(first_sentence)
    return patterns


def _fmt_survey_value(value: float | None, *, label: str) -> str:
    if not has_numeric_value(value):
        return data_not_found(label)
    number = float(value)
    if number == int(number):
        return str(int(number))
    return f"{number:.1f}"


def _fmt_survey_bw(actual: float | None, goal: float | None, *, label: str) -> str:
    if not has_numeric_value(actual) or not has_numeric_value(goal):
        return data_not_found(label)
    diff = float(actual) - float(goal)
    if diff < 0:
        if diff == int(diff):
            return f"({int(abs(diff))})"
        return f"({abs(diff):.1f})"
    if diff == int(diff):
        return str(int(diff))
    return f"{diff:.1f}"


def _fmt_score(value: float | None) -> str:
    if not has_numeric_value(value):
        return ""
    return f"{float(value):.2f}"


def _month_header(data: MprData) -> str:
    from mpr_data import MONTH_LABELS

    return MONTH_LABELS[data.month - 1].upper()


def _scorecard_month_columns(block, header_row: int | None, data: MprData) -> tuple[int | None, int | None]:
    if header_row is None:
        return None, None
    month_header = _month_header(data)
    mtd_col = ytd_col = None
    for col_idx in range(len(block.columns)):
        header = format_scorecard_cell(block.iat[header_row, col_idx]).strip().upper()
        if header == month_header:
            mtd_col = col_idx
        if header in {"YTD", "YE"}:
            ytd_col = col_idx
    return mtd_col, ytd_col


def _scorecard_scores(data: MprData) -> dict[str, tuple[float | None, float | None]]:
    scores: dict[str, tuple[float | None, float | None]] = {}
    block = load_system_scorecard(data, block="summary_2")
    if block.empty:
        return scores

    header_row = find_scorecard_header_row(block)
    mtd_col, ytd_col = _scorecard_month_columns(block, header_row, data)
    index = build_scorecard_row_index(block, header_row=header_row)

    for section_label, patterns in PEOPLE_SECTIONS.items():
        kpi_name = patterns[0]
        row_idx = index.get(f"{kpi_name.lower()}|score")
        if row_idx is None:
            continue
        mtd_score = ytd_score = None
        if mtd_col is not None:
            val = block.iat[row_idx, mtd_col]
            if format_scorecard_cell(val):
                mtd_score = float(val)
        if ytd_col is not None:
            val = block.iat[row_idx, ytd_col]
            if format_scorecard_cell(val):
                ytd_score = float(val)
        scores[section_label] = (mtd_score, ytd_score)
    return scores


def _lookup_survey_kpi(
    data: MprData,
    patterns: list[str],
    workbook: str,
    *,
    mtd: bool,
) -> tuple[float | None, float | None]:
    if mtd:
        result = lookup_kpi_exhaustive(data, patterns, month=data.month)
        return result.value.actual, result.value.goal
    actual, _ = lookup_ytd_exhaustive(data, patterns)
    goal = lookup_kpi_exhaustive(data, patterns, month=data.month).value.goal
    return actual, goal


def _clear_row_values(table, row_idx: int, columns: dict[str, int]) -> None:
    for col_idx in columns.values():
        set_cell_text_preserve(table.cell(row_idx, col_idx), "")


def fill_people_table(
    table,
    data: MprData,
    *,
    workbook: str,
    scorecards: dict[str, tuple[float | None, float | None]] | None = None,
) -> int:
    if not is_people_table(table):
        return 0

    layout = parse_people_table_layout(table)
    if not layout:
        return 0

    scorecards = scorecards if scorecards is not None else _scorecard_scores(data)
    filled = 0
    current_section: str | None = None

    for row_idx in range(len(table.rows)):
        label = table.cell(row_idx, 0).text.strip()
        if not label:
            continue
        normalized = _normalize_label(label)
        if normalized in PEOPLE_SECTIONS:
            current_section = normalized
        elif label.upper() in {"PEOPLE", "MTD", "YTD", "SCORE", "ACTUAL", "B/(W) GOAL"}:
            continue
        elif normalized.startswith("MTD") or normalized.startswith("YTD"):
            continue

        patterns = _kpi_patterns_for_row(label)
        if not patterns:
            continue

        mtd_actual, mtd_goal = _lookup_survey_kpi(data, patterns, workbook, mtd=True)
        ytd_actual, ytd_goal = _lookup_survey_kpi(data, patterns, workbook, mtd=False)

        metric_name = patterns[0]
        values = {
            "mtd_actual": _fmt_survey_value(mtd_actual, label=f"{metric_name} MTD Actual"),
            "mtd_bw": _fmt_survey_bw(mtd_actual, mtd_goal, label=f"{metric_name} MTD B/(W) Goal"),
            "ytd_actual": _fmt_survey_value(ytd_actual, label=f"{metric_name} YTD Actual"),
            "ytd_bw": _fmt_survey_bw(ytd_actual, ytd_goal, label=f"{metric_name} YTD B/(W) Goal"),
        }

        if normalized in PEOPLE_SECTIONS and normalized in scorecards:
            mtd_score, ytd_score = scorecards[normalized]
            values["score_mtd"] = _fmt_score(mtd_score)
            values["score_ytd"] = _fmt_score(ytd_score)
        elif normalized in PEOPLE_SECTIONS:
            values["score_mtd"] = ""
            values["score_ytd"] = ""
        else:
            section_scores = scorecards.get(current_section or "", (None, None))
            values["score_mtd"] = _fmt_score(section_scores[0])
            values["score_ytd"] = _fmt_score(section_scores[1])

        for metric, col_idx in layout.items():
            text = values.get(metric)
            if text is None:
                continue
            if text == "" and metric.startswith("score_"):
                continue
            set_cell_text_preserve(table.cell(row_idx, col_idx), text)
            if text not in {"", EMPTY_CELL}:
                filled += 1

    return filled


CHART_TO_PATTERNS = {
    "Leadership Engagement": PEOPLE_SECTIONS["LEADERSHIP ENGAGEMENT"],
    "Psychological Safety": PEOPLE_SECTIONS["PSYCHOLOGICAL SAFETY"],
    "Accountability": PEOPLE_SECTIONS["ACCOUNTABILITY"],
}


def chart_configs_from_element(element: dict) -> list[dict]:
    charts = element.get("charts")
    if charts:
        return charts
    return [{"title": title, "patterns": CHART_TO_PATTERNS[title]} for title in PEOPLE_CHART_ORDER]
