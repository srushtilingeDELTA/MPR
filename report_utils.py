"""Shared report period and config helpers."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import yaml
from dateutil.relativedelta import relativedelta

MONTH_FROM_SHEET = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def parse_bool(value, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in ("true", "yes", "1", "on"):
            return True
        if normalized in ("false", "no", "0", "off"):
            return False
    return default


def month_from_sheet_name(sheet_name: str) -> int | None:
    if not sheet_name:
        return None
    token = sheet_name.strip().split()[0].lower()
    return MONTH_FROM_SHEET.get(token)


def load_config(path: Path | None = None, base_dir: Path | None = None) -> dict:
    root = base_dir or Path(__file__).resolve().parent
    config_path = path or root / "config.yaml"
    if not config_path.exists():
        raise FileNotFoundError(f"Config not found: {config_path}")

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
    use_previous = parse_bool(report_cfg.get("use_previous_month"), default=True)

    if use_previous:
        today = date.today()
        prev = today.replace(day=1) - relativedelta(months=1)
        return prev.year, prev.month, prev.strftime("%B %Y")

    year = report_cfg.get("year")
    month = report_cfg.get("month")
    if month is None:
        month = month_from_sheet_name(str(excel_cfg.get("sheet_name", "")))
    if year is None:
        year = date.today().year

    year = int(year)
    month = int(month)
    return year, month, date(year, month, 1).strftime("%B %Y")