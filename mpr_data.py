"""Load and aggregate KPI data from MPR Excel workbooks."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date
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

    def _match_kpi_rows(self, df: pd.DataFrame, kpi_patterns: list[str], month: int | None = None) -> pd.DataFrame:
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

        if aggregation == "first":
            row = rows.iloc[0]
            return KpiValue(
                actual=float(row[c["actual"]]) if pd.notna(row.get(c["actual"])) else None,
                goal=float(row[goal_col]) if goal_col and pd.notna(row.get(goal_col)) else None,
                num=float(row[c["num"]]) if c.get("num") in rows.columns and pd.notna(row.get(c["num"])) else None,
                den=float(row[c["den"]]) if c.get("den") in rows.columns and pd.notna(row.get(c["den"])) else None,
            )

        num = den = actual = goal = None
        if c.get("num") in rows.columns and c.get("den") in rows.columns:
            num = rows[c["num"]].sum(min_count=1)
            den = rows[c["den"]].sum(min_count=1)
            if pd.notna(num) and pd.notna(den) and den != 0:
                actual = float(num / den)

        if actual is None and c["actual"] in rows.columns:
            mean_val = rows[c["actual"]].mean()
            actual = float(mean_val) if pd.notna(mean_val) else None

        if goal_col:
            goals = rows[goal_col].dropna()
            if not goals.empty:
                goal_val = goals.iloc[0] if aggregation == "first" else goals.mean()
                goal = float(goal_val) if pd.notna(goal_val) else None

        num_out = float(num) if pd.notna(num) else None
        den_out = float(den) if pd.notna(den) else None
        return KpiValue(actual=actual, goal=goal, num=num_out, den=den_out)

    def kpi_value(
        self,
        kpi_patterns: list[str],
        month: int | None = None,
        aggregation: str = "weighted",
        workbook: str | None = None,
    ) -> KpiValue:
        month = month or self.month
        wb = workbook or self._actuals_workbook()
        df = self.month_frame(month, workbook=wb)
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

    def monthly_series(
        self,
        kpi_patterns: list[str],
        through_month: int | None = None,
        aggregation: str = "weighted",
        workbook: str | None = None,
    ) -> list[float | None]:
        through_month = through_month or self.month
        values: list[float | None] = []
        for m in range(1, 13):
            if m <= through_month:
                kv = self.kpi_value(kpi_patterns, month=m, aggregation=aggregation, workbook=workbook)
                values.append(kv.actual)
            else:
                values.append(None)
        return values

    def ytd_value(
        self,
        kpi_patterns: list[str],
        aggregation: str = "weighted",
        workbook: str | None = None,
    ) -> float | None:
        wb = workbook or self._actuals_workbook()
        parts = [
            v
            for v in self.monthly_series(kpi_patterns, through_month=self.month, aggregation=aggregation, workbook=wb)
            if v is not None
        ]
        if not parts:
            return None
        if aggregation == "weighted":
            nums = dens = 0.0
            for m in range(1, self.month + 1):
                df = self.month_frame(m, workbook=wb)
                rows = self._match_kpi_rows(df, kpi_patterns, month=m)
                c = self.cols
                if c.get("num") in rows.columns and c.get("den") in rows.columns:
                    nums += rows[c["num"]].sum(min_count=1) or 0
                    dens += rows[c["den"]].sum(min_count=1) or 0
            if dens:
                return nums / dens
        return sum(parts) / len(parts)

    def report_month_label(self) -> str:
        return date(self.year, self.month, 1).strftime("%B %Y")

    def report_month_short(self) -> str:
        return date(self.year, self.month, 1).strftime("%b'%y")

    def report_output_title(self) -> str:
        return f"GSE MPR - {self.report_month_label()}"