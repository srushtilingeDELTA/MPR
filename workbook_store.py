"""Load and query multiple MPR Excel workbooks (actuals, workings, scorecards)."""

from __future__ import annotations

import io
import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


class WorkbookStore:
    """In-memory registry of Excel workbooks used by the MPR report."""

    def __init__(self, config: dict, base_dir: Path):
        self.config = config
        self.base_dir = base_dir
        self._files: dict[str, pd.ExcelFile] = {}
        self._sheet_cache: dict[tuple[str, str], pd.DataFrame] = {}

    def _workbook_defs(self) -> dict:
        defs = self.config.get("workbooks")
        if defs:
            return defs
        return {
            "actuals": {
                "path": self.config["excel"]["file_path"],
                "header_row": self.config["excel"].get("header_row", 0),
            }
        }

    def _bytes_for_path(self, rel_path: str) -> bytes:
        cached = self.config.get("_sharepoint_files", {}).get(rel_path)
        if cached:
            return cached

        from sharepoint_live import get_cached_file

        cached = get_cached_file(rel_path)
        if cached:
            return cached

        local = self.base_dir / rel_path
        if local.exists():
            return local.read_bytes()

        raise FileNotFoundError(
            f"Workbook not available: {rel_path}. "
            "Sync SharePoint files or save a local copy under data/."
        )

    def load(self, *, only: list[str] | None = None) -> None:
        for name, wb_cfg in self._workbook_defs().items():
            if only is not None and name not in only:
                continue
            rel_path = wb_cfg["path"]
            try:
                data = self._bytes_for_path(rel_path)
            except FileNotFoundError:
                if wb_cfg.get("optional"):
                    logger.warning("Optional workbook %r not found: %s", name, rel_path)
                    continue
                raise
            self._files[name] = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
            logger.info(
                "Loaded workbook %r (%s bytes, %s sheets)",
                name,
                len(data),
                len(self._files[name].sheet_names),
            )

    def sheet_names(self, workbook: str) -> list[str]:
        xl = self._files.get(workbook)
        return xl.sheet_names if xl else []

    def read_sheet(
        self,
        workbook: str,
        sheet_name: str,
        *,
        header_row: int | None = None,
        raw: bool = False,
    ) -> pd.DataFrame:
        cache_key = (workbook, sheet_name, header_row, raw)
        if cache_key in self._sheet_cache:
            return self._sheet_cache[cache_key].copy()

        xl = self._files.get(workbook)
        if xl is None:
            return pd.DataFrame()

        if sheet_name not in xl.sheet_names:
            logger.debug("Sheet %r not found in workbook %r", sheet_name, workbook)
            return pd.DataFrame()

        if raw:
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None, engine="openpyxl")
        else:
            wb_cfg = self._workbook_defs().get(workbook, {})
            header = wb_cfg.get("header_row", 0) if header_row is None else header_row
            df = pd.read_excel(xl, sheet_name=sheet_name, header=header, engine="openpyxl")
            df.columns = df.columns.astype(str).str.strip()
        self._sheet_cache[cache_key] = df
        return df.copy()

    def read_sheet_by_index(self, workbook: str, index: int, *, header_row: int | None = None) -> pd.DataFrame:
        names = self.sheet_names(workbook)
        if index < 0 or index >= len(names):
            return pd.DataFrame()
        return self.read_sheet(workbook, names[index], header_row=header_row)

    def cell_value(self, workbook: str, sheet_name: str, row: int, col: int) -> object:
        df = self.read_sheet(workbook, sheet_name, header_row=None)
        if df.empty or row >= len(df.index) or col >= len(df.columns):
            return None
        return df.iat[row, col]

    def range_values(self, workbook: str, sheet_name: str, start_row: int, start_col: int, end_row: int, end_col: int) -> list[list[object]]:
        df = self.read_sheet(workbook, sheet_name, header_row=None)
        if df.empty:
            return []
        rows: list[list[object]] = []
        for r in range(start_row, min(end_row + 1, len(df.index))):
            row_vals: list[object] = []
            for c in range(start_col, min(end_col + 1, len(df.columns))):
                row_vals.append(df.iat[r, c])
            rows.append(row_vals)
        return rows

    def open_worksheet(self, workbook: str, sheet_name: str):
        """Return an openpyxl worksheet (with formatting) for styled table export."""
        import openpyxl

        rel_path = self._workbook_defs()[workbook]["path"]
        data = self._bytes_for_path(rel_path)
        xl = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
        if sheet_name not in xl.sheetnames:
            raise KeyError(f"Sheet {sheet_name!r} not in workbook {workbook!r}")
        return xl[sheet_name]
