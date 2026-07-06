"""Quick check that config.yaml and required folders/files are set up."""

from __future__ import annotations

from pathlib import Path

import yaml

BASE_DIR = Path(__file__).resolve().parent.parent


def main() -> None:
    config_path = BASE_DIR / "config.yaml"
    print(f"Project folder: {BASE_DIR}")
    print(f"Config file:    {config_path} ({'found' if config_path.exists() else 'MISSING'})")

    if not config_path.exists():
        return

    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    excel_cfg = config.get("excel", {})
    for key in ("file_path", "sheet_name"):
        value = excel_cfg.get(key, "MISSING")
        print(f"  excel.{key}: {value}")

    report_cfg = config.get("report", {})
    print(
        f"  report: use_previous_month={report_cfg.get('use_previous_month')}, "
        f"year={report_cfg.get('year')}, month={report_cfg.get('month')}"
    )

    excel_path = BASE_DIR / excel_cfg.get("file_path", "")
    template_path = BASE_DIR / config.get("powerpoint", {}).get("template_path", "")
    print(f"Excel file:     {excel_path} ({'found' if excel_path.exists() else 'MISSING'})")
    print(f"Template:       {template_path} ({'found' if template_path.exists() else 'MISSING'})")


if __name__ == "__main__":
    main()