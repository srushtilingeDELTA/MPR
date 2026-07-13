"""Load and aggregate KPI data from MPR Excel workbooks."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from workbook_store import WorkbookStore

logger = logging.getLogger(__name__)

MONTH_SHEET_NAMES = {
    1: "Jan Actuals",
    2: "Feb Actuals",
    3: "Mar Actuals",
    4: "Apr Actuals",
    5: "May Actuals",
    6: "June Actuals",
    7: "July Actuals",
    8: "Aug Actuals",
    9: "Sept Actuals",
    10: "Oct Actuals",
    11: "Nov Actuals",
    12: "Dec Actuals",
}

MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_LABELS_UPPER = [m.upper() for m in MONTH_LABELS]


def _coerce_number(value) -> float | None:
    """Convert Excel/pandas cell values to float for KPI aggregation."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, (timedelta, pd.Timedelta)):
        # Excel duration cells often arrive as timedelta; use day-fraction scale.
        return value.total_seconds() / 86400.0
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return None

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


@dataclass
class KpiValue:
    actual: float | None = None
    goal: float | None = None
    num: float | None = None
    den: float | None = None


class MprData:
    """Extract KPI values from configured Excel workbooks."""

    def __init__(self, config: dict, base_dir: Path):
        self.config = config
        self.base_dir = base_dir
        self.cols = config["columns"]
        self.year, self.month, _ = self._report_period()
        self.store = WorkbookStore(config, base_dir)
        self._frames: dict[int, pd.DataFrame] = {}
        self._sheet_names: list[str] = []

    def _report_period(self) -> tuple[int, int, str]:
        from report_utils import get_report_period

        return get_report_period(self.config)

    @property
    def file_path(self) -> Path:
        return self.base_dir / self.config["excel"]["file_path"]

    def _excel_source(self) -> str:
        excel_cfg = self.config.get("excel", {})
        if excel_cfg.get("source"):
            return str(excel_cfg["source"]).lower()
        if self.config.get("sharepoint", {}).get("live_read"):
            return "sharepoint"
        return "local"

    def _actuals_workbook(self) -> str:
        return "actuals"

    def _normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.dropna(axis=1, how="all").copy()
        df.columns = df.columns.astype(str).str.strip()
        c = self.cols
        for col in (c["year"], c["month"], c["actual"], c.get("num"), c.get("den")):
            if col and col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        if c["kpi"] in df.columns:
            df[c["kpi"]] = df[c["kpi"]].astype(str).str.strip()
        if c.get("entity") in df.columns:
            df[c["entity"]] = df[c["entity"]].astype(str).str.strip()
        return df

    def _load_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
        excel_cfg = self.config["excel"]
        kwargs: dict = {"sheet_name": sheet_name}
        if "header_row" in excel_cfg:
            kwargs["header"] = excel_cfg["header_row"]
        return self._normalize(pd.read_excel(xl, **kwargs))

    def load(self) -> None:
        self.store.load()
        actuals = self._actuals_workbook()
        xl = self.store._files.get(actuals)
        if xl is None:
            raise FileNotFoundError(f"Primary workbook {actuals!r} was not loaded.")

        self._sheet_names = xl.sheet_names
        primary = self.config["excel"]["sheet_name"]
        if primary in self._sheet_names:
            self._frames[self.month] = self._load_sheet(xl, primary)

        for month_num, sheet in MONTH_SHEET_NAMES.items():
            if sheet in self._sheet_names and month_num not in self._frames:
                self._frames[month_num] = self._load_sheet(xl, sheet)

    def month_frame(self, month: int, workbook: str | None = None) -> pd.DataFrame:
        if workbook and workbook != self._actuals_workbook():
            sheet = MONTH_SHEET_NAMES.get(month)
            if not sheet:
                return pd.DataFrame()
            df = self.store.read_sheet(workbook, sheet)
            return self._normalize(df) if not df.empty else df
        return self._frames.get(month, pd.DataFrame())

    def read_sheet(self, workbook: str, sheet_name: str) -> pd.DataFrame:
        return self.store.read_sheet(workbook, sheet_name)

    def sheet_names(self, workbook: str) -> list[str]:
        return self.store.sheet_names(workbook)

    def _goal_column(self, df: pd.DataFrame) -> str | None:
        for candidate in ("Goal", "Plan", "Target", "GOAL"):
            if candidate in df.columns:
                return candidate
        return None

    def _match_kpi_rows(
        self,
        df: pd.DataFrame,
        kpi_patterns: list[str],
        month: int | None = None,
        *,
        year: int | None = None,
    ) -> pd.DataFrame:
        if df.empty or self.cols["kpi"] not in df.columns:
            return df.iloc[0:0]

        col = self.cols["kpi"]
        mask = pd.Series(False, index=df.index)
        for pattern in kpi_patterns:
            exact = df[col].str.strip().str.casefold() == pattern.strip().casefold()
            mask |= exact
            if not exact.any():
                mask |= df[col].str.contains(pattern, case=False, na=False, regex=False)
        rows = df.loc[mask].copy()
        c = self.cols
        target_year = year if year is not None else self.year
        if c["year"] in rows.columns:
            rows = rows.loc[rows[c["year"]] == target_year]
        if month and c["month"] in rows.columns:
            rows = rows.loc[rows[c["month"]] == month]
        return rows

    def _system_rows(self, rows: pd.DataFrame) -> pd.DataFrame:
        entity_col = self.cols.get("entity")
        if entity_col and entity_col in rows.columns:
            system_rows = rows[rows[entity_col].str.contains("system", case=False, na=False)]
            if not system_rows.empty:
                return system_rows
        return rows

    def _entity_rows(self, rows: pd.DataFrame, entity_pattern: str) -> pd.DataFrame:
        entity_col = self.cols.get("entity")
        if entity_col and entity_col in rows.columns:
            matched = rows[rows[entity_col].str.contains(entity_pattern, case=False, na=False, regex=False)]
            if not matched.empty:
                return matched
        return rows.iloc[0:0]

    def _rows_for_kpi(
        self,
        kpi_patterns: list[str],
        month: int,
        *,
        year: int | None = None,
        workbook: str | None = None,
        entity_pattern: str | None = None,
        system_only: bool = True,
    ) -> pd.DataFrame:
        wb = workbook or self._actuals_workbook()
        df = self.month_frame(month, workbook=wb)
        if df.empty:
            return df.iloc[0:0]

        rows = self._match_kpi_rows(df, kpi_patterns, month=month, year=year)
        if rows.empty:
            return rows

        if entity_pattern:
            rows = self._entity_rows(rows, entity_pattern)
        elif system_only:
            system_rows = self._system_rows(rows)
            if not system_rows.empty:
                rows = system_rows
        return rows

    def _aggregate_rows(self, rows: pd.DataFrame, aggregation: str = "weighted") -> KpiValue:
        if rows.empty:
            return KpiValue()

        c = self.cols
        goal_col = self._goal_column(rows)

        if aggregation == "first":
            row = rows.iloc[0]
            return KpiValue(
                actual=_coerce_number(row.get(c["actual"])),
                goal=_coerce_number(row.get(goal_col)) if goal_col else None,
                num=_coerce_number(row.get(c["num"])) if c.get("num") in rows.columns else None,
                den=_coerce_number(row.get(c["den"])) if c.get("den") in rows.columns else None,
            )

        num = den = actual = goal = None

        if aggregation == "sum":
            if c["actual"] in rows.columns:
                total = rows[c["actual"]].sum(min_count=1)
                actual = _coerce_number(total)
            goal = None
            if goal_col:
                goals = rows[goal_col].dropna()
                if not goals.empty:
                    goal = _coerce_number(goals.iloc[0])
            return KpiValue(actual=actual, goal=goal)

        if c["actual"] in rows.columns:
            mean_val = rows[c["actual"]].mean()
            actual = _coerce_number(mean_val)

        if actual is None and c.get("num") in rows.columns and c.get("den") in rows.columns:
            num = rows[c["num"]].sum(min_count=1)
            den = rows[c["den"]].sum(min_count=1)
            if pd.notna(num) and pd.notna(den) and den != 0:
                actual = _coerce_number(num / den)

        if goal_col:
            goals = rows[goal_col].dropna()
            if not goals.empty:
                goal_val = goals.iloc[0] if aggregation == "first" else goals.mean()
                goal = _coerce_number(goal_val)

        num_out = _coerce_number(num)
        den_out = _coerce_number(den)
        return KpiValue(actual=actual, goal=goal, num=num_out, den=den_out)

    def kpi_value(
        self,
        kpi_patterns: list[str],
        month: int | None = None,
        aggregation: str = "weighted",
        workbook: str | None = None,
        *,
        year: int | None = None,
        entity_pattern: str | None = None,
        system_only: bool = True,
    ) -> KpiValue:
        month = month or self.month
        rows = self._rows_for_kpi(
            kpi_patterns,
            month,
            year=year,
            workbook=workbook,
            entity_pattern=entity_pattern,
            system_only=system_only,
        )
        if rows.empty:
            return KpiValue()
        return self._aggregate_rows(rows, aggregation=aggregation)

    def monthly_series(
        self,
        kpi_patterns: list[str],
        through_month: int | None = None,
        aggregation: str = "weighted",
        workbook: str | None = None,
        *,
        year: int | None = None,
        entity_pattern: str | None = None,
        system_only: bool = True,
    ) -> list[float | None]:
        through_month = through_month or self.month
        target_year = year if year is not None else self.year
        values: list[float | None] = []
        for m in range(1, 13):
            if m <= through_month:
                kv = self.kpi_value(
                    kpi_patterns,
                    month=m,
                    aggregation=aggregation,
                    workbook=workbook,
                    year=target_year,
                    entity_pattern=entity_pattern,
                    system_only=system_only,
                )
                values.append(kv.actual)
            else:
                values.append(None)
        return values

    def ytd_value(
        self,
        kpi_patterns: list[str],
        aggregation: str = "weighted",
        workbook: str | None = None,
        *,
        year: int | None = None,
        entity_pattern: str | None = None,
        system_only: bool = True,
    ) -> float | None:
        wb = workbook or self._actuals_workbook()
        target_year = year if year is not None else self.year
        parts = [
            v
            for v in self.monthly_series(
                kpi_patterns,
                through_month=self.month,
                aggregation=aggregation,
                workbook=wb,
                year=target_year,
                entity_pattern=entity_pattern,
                system_only=system_only,
            )
            if v is not None
        ]
        if not parts:
            return None
        if aggregation == "weighted":
            nums = dens = 0.0
            c = self.cols
            for m in range(1, self.month + 1):
                rows = self._rows_for_kpi(
                    kpi_patterns,
                    m,
                    year=target_year,
                    workbook=wb,
                    entity_pattern=entity_pattern,
                    system_only=system_only,
                )
                if c.get("num") in rows.columns and c.get("den") in rows.columns:
                    nums += rows[c["num"]].sum(min_count=1) or 0
                    dens += rows[c["den"]].sum(min_count=1) or 0
            if dens:
                return nums / dens
        return sum(parts) / len(parts)

    def ytd_sum(
        self,
        kpi_patterns: list[str],
        workbook: str | None = None,
        *,
        year: int | None = None,
        entity_pattern: str | None = None,
        system_only: bool = False,
    ) -> float | None:
        total = 0.0
        found = False
        target_year = year if year is not None else self.year
        for month in range(1, self.month + 1):
            kv = self.kpi_value(
                kpi_patterns,
                month=month,
                aggregation="sum",
                workbook=workbook,
                year=target_year,
                entity_pattern=entity_pattern,
                system_only=system_only,
            )
            if kv.actual is not None:
                total += float(kv.actual)
                found = True
        return total if found else None

    def prior_year_monthly_series(
        self,
        kpi_patterns: list[str],
        *,
        years_back: int = 1,
        through_month: int | None = None,
        workbook: str | None = None,
    ) -> list[float | None]:
        return self.monthly_series(
            kpi_patterns,
            through_month=through_month or self.month,
            workbook=workbook,
            year=self.year - years_back,
        )

    def report_month_label(self) -> str:
        return date(self.year, self.month, 1).strftime("%B %Y")

    def report_month_short(self) -> str:
        return date(self.year, self.month, 1).strftime("%b'%y")

    def report_output_title(self) -> str:
        return f"GSE MPR - {self.report_month_label()}"
