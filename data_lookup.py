"""Search all loaded workbooks before reporting data as missing."""

from __future__ import annotations

import logging
from dataclasses import dataclass

from mpr_data import KpiValue, MONTH_SHEET_NAMES

logger = logging.getLogger(__name__)

WORKBOOK_SEARCH_ORDER = ("actuals", "workings", "scorecards")


@dataclass
class KpiLookupResult:
    value: KpiValue
    workbook: str | None = None
    sheet: str | None = None

    @property
    def found(self) -> bool:
        return self.value.actual is not None or self.value.goal is not None


def _available_workbooks(data) -> list[str]:
    loaded = list(data.store._files.keys())
    ordered = [wb for wb in WORKBOOK_SEARCH_ORDER if wb in loaded]
    for wb in loaded:
        if wb not in ordered:
            ordered.append(wb)
    return ordered


def _has_kpi_schema(data, df) -> bool:
    """Only search sheets that look like monthly actuals tables."""
    cols = data.cols
    required = (cols.get("kpi"), cols.get("actual"), cols.get("year"), cols.get("month"))
    return all(col and col in df.columns for col in required)


def _sheet_search_order(data, workbook: str, month: int) -> list[str]:
    names = data.store.sheet_names(workbook)
    if not names:
        return []
    preferred = MONTH_SHEET_NAMES.get(month)
    ordered: list[str] = []
    if preferred and preferred in names:
        ordered.append(preferred)
    for sheet in names:
        if sheet not in ordered:
            ordered.append(sheet)
    return ordered


def lookup_kpi_exhaustive(
    data,
    kpi_patterns: list[str],
    *,
    month: int | None = None,
    year: int | None = None,
    entity_pattern: str | None = None,
    aggregation: str = "weighted",
) -> KpiLookupResult:
    """Search every loaded SharePoint/local workbook and sheet for a KPI."""
    month = month or data.month
    strategies = (
        {"system_only": True, "aggregation": aggregation},
        {"system_only": False, "aggregation": aggregation},
        {"system_only": True, "aggregation": "first"},
        {"system_only": False, "aggregation": "first"},
    )

    for workbook in _available_workbooks(data):
        for sheet in _sheet_search_order(data, workbook, month):
            df = data.store.read_sheet(workbook, sheet)
            if df.empty:
                continue
            df = data._normalize(df)
            if not _has_kpi_schema(data, df):
                continue
            rows = data._match_kpi_rows(df, kpi_patterns, month=month, year=year)
            if rows.empty:
                continue
            for strategy in strategies:
                scoped = rows
                if entity_pattern:
                    scoped = data._entity_rows(rows, entity_pattern)
                elif strategy["system_only"]:
                    scoped = data._system_rows(rows)
                    if scoped.empty:
                        scoped = rows
                if scoped.empty:
                    continue
                value = data._aggregate_rows(scoped, aggregation=strategy["aggregation"])
                if value.actual is not None or value.goal is not None:
                    logger.debug(
                        "Found KPI %r in %s / %s (strategy=%s)",
                        kpi_patterns[0],
                        workbook,
                        sheet,
                        strategy,
                    )
                    return KpiLookupResult(value=value, workbook=workbook, sheet=sheet)

    logger.debug("KPI not found after exhaustive search: %r", kpi_patterns)
    return KpiLookupResult(value=KpiValue())


def lookup_monthly_series_exhaustive(
    data,
    kpi_patterns: list[str],
    *,
    through_month: int | None = None,
    year: int | None = None,
    entity_pattern: str | None = None,
    aggregation: str = "first",
) -> tuple[list[float | None], KpiLookupResult]:
    """Build a monthly series, searching all workbooks if the primary lookup is empty."""
    through_month = through_month or data.month
    values: list[float | None] = []
    source = KpiLookupResult(value=KpiValue())

    for month_num in range(1, 13):
        if month_num <= through_month:
            result = lookup_kpi_exhaustive(
                data,
                kpi_patterns,
                month=month_num,
                year=year,
                entity_pattern=entity_pattern,
                aggregation=aggregation,
            )
            values.append(result.value.actual)
            if result.found and source.workbook is None:
                source = result
        else:
            values.append(None)

    if source.found:
        return values, source

    # Fall back to the original month-frame path (cached report-month sheet).
    fallback = [
        data.kpi_value(
            kpi_patterns,
            month=m,
            year=year,
            entity_pattern=entity_pattern,
            aggregation=aggregation,
            system_only=False,
        ).actual
        if m <= through_month
        else None
        for m in range(1, 13)
    ]
    if any(v is not None for v in fallback):
        return fallback, KpiLookupResult(
            value=data.kpi_value(
                kpi_patterns,
                month=through_month,
                year=year,
                entity_pattern=entity_pattern,
                aggregation=aggregation,
                system_only=False,
            ),
            workbook="actuals",
            sheet=None,
        )
    return values, source


def lookup_ytd_exhaustive(
    data,
    kpi_patterns: list[str],
    *,
    year: int | None = None,
    entity_pattern: str | None = None,
    aggregation: str = "weighted",
) -> tuple[float | None, KpiLookupResult]:
    """YTD from weighted num/den when available; fall back to monthly series."""
    ytd = data.ytd_value(
        kpi_patterns,
        year=year,
        entity_pattern=entity_pattern,
        aggregation=aggregation,
        system_only=entity_pattern is None,
    )
    if ytd is not None:
        return ytd, KpiLookupResult(
            value=KpiValue(actual=ytd),
            workbook="actuals",
            sheet=None,
        )

    series, source = lookup_monthly_series_exhaustive(
        data,
        kpi_patterns,
        year=year,
        entity_pattern=entity_pattern,
        aggregation="first" if aggregation == "weighted" else aggregation,
    )
    parts = [v for v in series if v is not None]
    if not parts:
        return None, KpiLookupResult(value=KpiValue())
    if aggregation == "sum":
        return float(sum(parts)), source
    return float(sum(parts) / len(parts)), source
