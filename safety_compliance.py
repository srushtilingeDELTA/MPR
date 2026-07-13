"""Fill EA Compliance / ASAP Reporting tables from MPR Actuals."""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING

from data_lookup import lookup_kpi_exhaustive, lookup_ytd_exhaustive
from narrative_boxes import clear_manual_narrative_boxes
from ppt_format import set_cell_text_preserve
from ppt_missing import EMPTY_CELL, data_not_found, has_numeric_value
from scorecard_style import apply_compliance_table_styles

if TYPE_CHECKING:
    from mpr_data import MprData

logger = logging.getLogger(__name__)

ACTUALS_SOURCE = "MPR Actuals and Goals_v2.xlsx"

# Template row label -> Excel Entity search patterns (first match wins).
REGION_ENTITY_PATTERNS: list[tuple[str, list[str]]] = [
    ("ATL", ["ATL (M)", "ATL(M)"]),
    ("BOS", ["BOS (M)", "BOS(M)"]),
    ("DTW", ["DTW (M)", "DTW(M)"]),
    ("FLORIDA", ["FLORIDA", "FL (M)", "FLORIDA (M)"]),
    ("LAX", ["LAX (M)", "LAX(M)"]),
    ("MSP", ["MSP (M)", "MSP(M)"]),
    ("NY", ["NY (M)", "NY(M)", "NYC (M)"]),
    ("SLC", ["SLC (M)", "SLC(M)"]),
    ("MOTORIZED", ["MOTORIZED", "Motorized"]),
    ("ATL (S)", ["ATL (S)", "ATL(S)"]),
    ("DTW(S)", ["DTW (S)", "DTW(S)"]),
    ("JFK (S)", ["JFK (S)", "JFK(S)"]),
    ("LAX (S)", ["LAX (S)", "LAX(S)"]),
    ("LGA (S)", ["LGA (S)", "LGA(S)"]),
    ("MSP (S)", ["MSP (S)", "MSP(S)"]),
    ("STATIONARY", ["STATIONARY", "Stationary"]),
    ("SYSTEM", ["System", "SYSTEM"]),
]

METRIC_KEYS = ("mtd_actual", "mtd_bw", "ytd_actual", "ytd_bw", "score_mtd", "score_ytd")


def _normalize_region_label(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip().upper())


def _entity_patterns_for_row(row_label: str) -> list[str] | None:
    normalized = _normalize_region_label(row_label)
    for table_label, patterns in REGION_ENTITY_PATTERNS:
        if normalized == table_label or normalized.replace(" ", "") == table_label.replace(" ", ""):
            return patterns
    return None


def _table_text(table) -> str:
    return " ".join(
        table.cell(row, col).text.strip().lower()
        for row in range(len(table.rows))
        for col in range(len(table.columns))
    )


def is_safety_compliance_table(table) -> bool:
    text = _table_text(table)
    if "ea compliance" in text or "asap reporting" in text:
        return True
    if "regions" in text and ("asap" in text or "compliance" in text):
        return True
    labels = {_normalize_region_label(table.cell(row, 0).text) for row in range(len(table.rows))}
    return "ATL" in labels and ("SYSTEM" in labels or "MOTORIZED" in labels)


def _section_for_column(table, col_idx: int) -> str | None:
    for row_idx in range(min(3, len(table.rows))):
        for scan_col in range(col_idx, -1, -1):
            header = table.cell(row_idx, scan_col).text.strip().lower()
            if "asap" in header:
                return "asap"
            if "ea compliance" in header or header == "ea compliance":
                return "eac"
    return None


def _column_label(table, col_idx: int) -> str:
    parts = []
    for row_idx in range(min(3, len(table.rows))):
        text = table.cell(row_idx, col_idx).text.strip().lower()
        if text:
            parts.append(text)
    return " | ".join(parts)


def _column_header_context(table, col_idx: int) -> str:
    parts = []
    for row_idx in range(min(3, len(table.rows))):
        text = table.cell(row_idx, col_idx).text.strip().lower()
        if text:
            parts.append(text)
    for scan_col in range(col_idx, -1, -1):
        group = table.cell(1, scan_col).text.strip().lower() if len(table.rows) > 1 else ""
        if group in {"mtd", "ytd", "score"}:
            parts.insert(0, group)
            break
    return " | ".join(parts)


def _metric_from_label(label: str) -> str | None:
    if "score" in label:
        if "ytd" in label:
            return "score_ytd"
        return "score_mtd"
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


def parse_compliance_table_layout(table) -> dict[str, dict[str, int]]:
    """Map EAC/ASAP metric names to column indices."""
    layout: dict[str, dict[str, int]] = {"eac": {}, "asap": {}}
    for col_idx in range(len(table.columns)):
        section = _section_for_column(table, col_idx)
        if section is None:
            continue
        metric = _metric_from_label(_column_header_context(table, col_idx))
        if metric:
            layout[section][metric] = col_idx

    # Single-section tables: infer section from table title text.
    if not layout["eac"] and not layout["asap"]:
        return layout
    if not layout["eac"] or not layout["asap"]:
        text = _table_text(table)
        only = layout["eac"] or layout["asap"]
        if "asap" in text and not layout["asap"]:
            layout["asap"] = dict(only)
            layout["eac"] = {}
        elif ("ea compliance" in text or "eac" in text) and not layout["eac"]:
            layout["eac"] = dict(only)
            layout["asap"] = {}
    return layout


def _fmt_pct(value: float | None, *, label: str) -> str:
    if not has_numeric_value(value):
        return data_not_found(label)
    return f"{float(value):.1f}%"


def _fmt_count(value: float | None, *, label: str) -> str:
    if not has_numeric_value(value):
        return EMPTY_CELL
    number = float(value)
    if number == int(number):
        return str(int(number))
    return f"{number:.1f}"


def _fmt_bw(actual: float | None, goal: float | None, *, label: str, as_percent: bool) -> str:
    if not has_numeric_value(actual) or not has_numeric_value(goal):
        return data_not_found(label) if as_percent else EMPTY_CELL
    diff = float(actual) - float(goal)
    if as_percent:
        return f"{diff:.1f}%"
    if diff == int(diff):
        return str(int(diff))
    return f"{diff:.1f}"


def _region_data_rows(table) -> list[tuple[int, str, list[str]]]:
    rows: list[tuple[int, str, list[str]]] = []
    for row_idx in range(len(table.rows)):
        label = table.cell(row_idx, 0).text.strip()
        if not label:
            continue
        patterns = _entity_patterns_for_row(label)
        if patterns:
            rows.append((row_idx, label, patterns))
    return rows


def _lookup_kpi(
    data: MprData,
    kpi_patterns: list[str],
    entity_patterns: list[str],
    workbook: str,
    *,
    mtd: bool,
) -> tuple[float | None, float | None]:
    for entity_pattern in entity_patterns:
        entity = None if entity_pattern.strip().lower() == "system" else entity_pattern
        if mtd:
            result = lookup_kpi_exhaustive(
                data,
                kpi_patterns,
                month=data.month,
                entity_pattern=entity,
            )
            if result.found:
                return result.value.actual, result.value.goal
        else:
            actual, _ = lookup_ytd_exhaustive(data, kpi_patterns, entity_pattern=entity)
            goal = lookup_kpi_exhaustive(
                data,
                kpi_patterns,
                month=data.month,
                entity_pattern=entity,
            ).value.goal
            if actual is not None or goal is not None:
                return actual, goal
    return None, None


def _clear_region_row_values(table, row_idx: int, columns: dict[str, int]) -> None:
    for col_idx in columns.values():
        if col_idx > 0:
            set_cell_text_preserve(table.cell(row_idx, col_idx), "")


def fill_compliance_table(
    table,
    data: MprData,
    *,
    workbook: str,
    eac_patterns: list[str],
    asap_patterns: list[str],
) -> int:
    """Fill EA Compliance / ASAP cells from Excel; return number of cells written."""
    if not is_safety_compliance_table(table):
        return 0

    layout = parse_compliance_table_layout(table)
    if not layout["eac"] and not layout["asap"]:
        return 0

    region_rows = _region_data_rows(table)
    if not region_rows:
        return 0

    filled = 0
    section_kpis = {"eac": (eac_patterns, True), "asap": (asap_patterns, False)}

    for row_idx, row_label, entity_patterns in region_rows:
        for section, columns in layout.items():
            if not columns:
                continue
            kpi_patterns, as_percent = section_kpis[section]

            mtd_actual, mtd_goal = _lookup_kpi(
                data, kpi_patterns, entity_patterns, workbook, mtd=True
            )
            ytd_actual, ytd_goal = _lookup_kpi(
                data, kpi_patterns, entity_patterns, workbook, mtd=False
            )

            metric_values = {
                "mtd_actual": _fmt_pct(mtd_actual, label=f"{row_label} {section.upper()} MTD Actual")
                if as_percent
                else _fmt_count(mtd_actual, label=f"{row_label} {section.upper()} MTD Actual"),
                "mtd_bw": _fmt_bw(
                    mtd_actual,
                    mtd_goal,
                    label=f"{row_label} {section.upper()} MTD B/(W) Goal",
                    as_percent=as_percent,
                ),
                "ytd_actual": _fmt_pct(ytd_actual, label=f"{row_label} {section.upper()} YTD Actual")
                if as_percent
                else _fmt_count(ytd_actual, label=f"{row_label} {section.upper()} YTD Actual"),
                "ytd_bw": _fmt_bw(
                    ytd_actual,
                    ytd_goal,
                    label=f"{row_label} {section.upper()} YTD B/(W) Goal",
                    as_percent=as_percent,
                ),
            }

            for metric, col_idx in columns.items():
                if metric.startswith("score_"):
                    continue
                text = metric_values.get(metric)
                if text is None:
                    continue
                set_cell_text_preserve(table.cell(row_idx, col_idx), text)
                if text != EMPTY_CELL:
                    filled += 1

    return filled


EA_ASAP_TABLE_HEADERS = [
    [
        "Regions",
        "EA COMPLIANCE",
        "",
        "",
        "",
        "",
        "",
        "ASAP REPORTING",
        "",
        "",
        "",
        "",
        "",
    ],
    [
        "",
        "MTD",
        "",
        "YTD",
        "",
        "Score",
        "",
        "MTD",
        "",
        "YTD",
        "",
        "Score",
        "",
    ],
    [
        "",
        "Actual",
        "B/(W) Goal",
        "Actual",
        "B/(W) Goal",
        "MTD",
        "YTD",
        "Actual",
        "B/(W) Goal",
        "Actual",
        "B/(W) Goal",
        "MTD",
        "YTD",
    ],
]


def _compliance_region_labels() -> list[str]:
    return [label for label, _ in REGION_ENTITY_PATTERNS]


def install_compliance_table(slide) -> object | None:
    """Replace screenshot/OLE on slide with a native EA/ASAP table shell."""
    from picture_replace import find_largest_data_placeholder, replace_placeholder_with_table

    placeholder = find_largest_data_placeholder(slide)
    if placeholder is None:
        return None

    rows_data = EA_ASAP_TABLE_HEADERS + [[label] + [""] * 12 for label in _compliance_region_labels()]
    table = replace_placeholder_with_table(slide, placeholder, len(rows_data), len(rows_data[0]))
    for r, row in enumerate(rows_data):
        for c, value in enumerate(row):
            set_cell_text_preserve(table.cell(r, c), value)
    apply_compliance_table_styles(table)
    logger.info("Installed native EA/ASAP table (%sx%s)", len(rows_data), len(rows_data[0]))
    return table


def fill_compliance_slide(slide, data: MprData, *, workbook: str, eac_patterns: list[str], asap_patterns: list[str]) -> int:
    """Ensure EA/ASAP table exists on slide, then fill from actuals."""
    from picture_replace import remove_data_placeholders

    filled_total = 0
    tables = [shape.table for shape in slide.shapes if shape.has_table]
    if not tables:
        installed = install_compliance_table(slide)
        if installed is not None:
            tables = [installed]
    else:
        remove_data_placeholders(slide)

    for table in tables:
        filled_total += fill_compliance_table(
            table,
            data,
            workbook=workbook,
            eac_patterns=eac_patterns,
            asap_patterns=asap_patterns,
        )
        apply_compliance_table_styles(table)
    clear_manual_narrative_boxes(slide)
    return filled_total
