"""Shared report period and config helpers."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import yaml
from dateutil.relativedelta import relativedelta

MONTH_FROM_SHEET = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}


def parse_bool(value, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        n = value.strip().lower()
        if n in ("true", "yes", "1", "on"):
            return True
        if n in ("false", "no", "0", "off"):
            return False
    return default


def month_from_sheet_name(sheet_name: str) -> int | None:
    if not sheet_name:
        return None
    return MONTH_FROM_SHEET.get(sheet_name.strip().split()[0].lower())


def load_config(path: Path | None = None, base_dir: Path | None = None) -> dict:
    root = base_dir or Path(__file__).resolve().parent
    config_path = path or root / "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}
    excel_cfg = config.get("excel")
    if not isinstance(excel_cfg, dict):
        raise ValueError("config.yaml must include an 'excel:' section.")
    for key in ("file_path", "sheet_name"):
        if key not in excel_cfg:
            raise ValueError(f"config.yaml excel section missing '{key}'")
    return config


def get_report_period(config: dict) -> tuple[int, int, str]:
    report_cfg = config.get("report", {})
    excel_cfg = config.get("excel", {})
    if parse_bool(report_cfg.get("use_previous_month"), default=True):
        today = date.today()
        prev = today.replace(day=1) - relativedelta(months=1)
        return prev.year, prev.month, prev.strftime("%B %Y")
    year = int(report_cfg.get("year") or date.today().year)
    month = report_cfg.get("month") or month_from_sheet_name(str(excel_cfg.get("sheet_name", "")))
    month = int(month)
    return year, month, date(year, month, 1).strftime("%B %Y")