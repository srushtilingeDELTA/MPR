"""Load and aggregate KPI data from MPR Excel workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

import pandas as pd

MONTH_SHEET_NAMES = {
    1: "Jan Actuals", 2: "Feb Actuals", 3: "Mar Actuals", 4: "Apr Actuals",
    5: "May Actuals", 6: "June Actuals", 7: "July Actuals", 8: "Aug Actuals",
    9: "Sept Actuals", 10: "Oct Actuals", 11: "Nov Actuals", 12: "Dec Actuals",
}

MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@dataclass
class KpiValue:
    actual: float | None = None
    goal: float | None = None
    num: float | None = None
    den: float | None = None


class MprData:
    def __init__(self, config: dict, base_dir: Path):
        self.config = config
        self.base_dir = base_dir
        self.cols = config["columns"]
        self.year, self.month, _ = self._report_period()
        self._frames: dict[int, pd.DataFrame] = {}

    def _report_period(self) -> tuple[int, int, str]:
        from report_utils import get_report_period
        return get_report_period(self.config)

    @property
    def file_path(self) -> Path:
        return self.base_dir / self.config["excel"]["file_path"]

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

    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        excel_cfg = self.config["excel"]
        kwargs = {"sheet_name": sheet_name, "engine": "openpyxl"}
        if "header_row" in excel_cfg:
            kwargs["header"] = excel_cfg["header_row"]
        return self._normalize(pd.read_excel(self.file_path, **kwargs))

    def load(self) -> None:
        xl = pd.ExcelFile(self.file_path, engine="openpyxl")
        primary = self.config["excel"]["sheet_name"]
        if primary in xl.sheet_names:
            self._frames[self.month] = self._load_sheet(primary)
        for month_num, sheet in MONTH_SHEET_NAMES.items():
            if sheet in xl.sheet_names and month_num not in self._frames:
                self._frames[month_num] = self._load_sheet(sheet)

    def month_frame(self, month: int) -> pd.DataFrame:
        return self._frames.get(month, pd.DataFrame())

    def _goal_column(self, df: pd.DataFrame) -> str | None:
        for candidate in ("Goal", "Plan", "Target", "GOAL"):
            if candidate in df.columns:
                return candidate
        return None

    def _match_kpi_rows(self, df: pd.DataFrame, kpi_patterns: list[str], month: int | None = None) -> pd.DataFrame:
        if df.empty or self.cols["kpi"] not in df.columns:
            return df.iloc[0:0]
        col = self.cols["kpi"]
        mask = pd.Series(False, index=df.index)
        for pattern in kpi_patterns:
            mask |= df[col].str.contains(pattern, case=False, na=False, regex=False)
        rows = df.loc[mask].copy()
        c = self.cols
        if c["year"] in rows.columns:
            rows = rows.loc[rows[c["year"]] == self.year]
        if month and c["month"] in rows.columns:
            rows = rows.loc[rows[c["month"]] == month]
        return rows

    def _aggregate_rows(self, rows: pd.DataFrame, aggregation: str = "weighted") -> KpiValue:
        if rows.empty:
            return KpiValue()
        c = self.cols
        goal_col = self._goal_column(rows)
        num = den = actual = goal = None
        if c.get("num") in rows.columns and c.get("den") in rows.columns:
            num = rows[c["num"]].sum(min_count=1)
            den = rows[c["den"]].sum(min_count=1)
            if pd.notna(num) and pd.notna(den) and den != 0:
                actual = float(num / den)
        if actual is None and c["actual"] in rows.columns:
            actual = float(rows[c["actual"]].mean())
        if goal_col:
            goals = rows[goal_col].dropna()
            if not goals.empty:
                goal = float(goals.iloc[0])
        return KpiValue(actual=actual, goal=goal)

    def kpi_value(self, kpi_patterns: list[str], month: int | None = None, aggregation: str = "weighted") -> KpiValue:
        month = month or self.month
        df = self.month_frame(month)
        if df.empty:
            return KpiValue()
        rows = self._match_kpi_rows(df, kpi_patterns, month=month)
        if rows.empty:
            return KpiValue()
        if self.cols.get("entity") in rows.columns:
            system_rows = rows[rows[self.cols["entity"]].str.contains("system", case=False, na=False)]
            if not system_rows.empty:
                rows = system_rows
        return self._aggregate_rows(rows, aggregation=aggregation)

    def monthly_series(self, kpi_patterns: list[str], through_month: int | None = None) -> list[float | None]:
        through_month = through_month or self.month
        return [
            self.kpi_value(kpi_patterns, month=m).actual if m <= through_month else None
            for m in range(1, 13)
        ]

    def ytd_value(self, kpi_patterns: list[str]) -> float | None:
        nums = dens = 0.0
        c = self.cols
        for m in range(1, self.month + 1):
            rows = self._match_kpi_rows(self.month_frame(m), kpi_patterns, month=m)
            if c.get("num") in rows.columns and c.get("den") in rows.columns:
                nums += rows[c["num"]].sum(min_count=1) or 0
                dens += rows[c["den"]].sum(min_count=1) or 0
        if dens:
            return nums / dens
        parts = [v for v in self.monthly_series(kpi_patterns) if v is not None]
        return sum(parts) / len(parts) if parts else None

    def report_month_label(self) -> str:
        return date(self.year, self.month, 1).strftime("%B %Y")

    def report_month_short(self) -> str:
        return date(self.year, self.month, 1).strftime("%b'%y")