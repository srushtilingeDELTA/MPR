"""Quick check that config.yaml, main.py, and required files are set up."""

from __future__ import annotations

from pathlib import Path

import yaml

BASE_DIR = Path(__file__).resolve().parent.parent
EXPECTED_MAIN_MARKERS = ("SCRIPT_VERSION", "parse_bool", "Report settings")


def check_main_py(main_path: Path) -> bool:
    if not main_path.exists():
        print(f"main.py:        MISSING at {main_path}")
        return False

    text = main_path.read_text(encoding="utf-8")
    line_count = text.count("\n") + 1
    ok = all(marker in text for marker in EXPECTED_MAIN_MARKERS)
    status = "OK (latest)" if ok else "OUT OF DATE — run update.ps1"
    print(f"main.py:        {status} ({line_count} lines)")
    if not ok:
        print("  Expected markers in main.py:", ", ".join(EXPECTED_MAIN_MARKERS))
    return ok


def main() -> None:
    config_path = BASE_DIR / "config.yaml"
    main_path = BASE_DIR / "main.py"

    print(f"Project folder: {BASE_DIR}")
    print(f"Config file:    {config_path} ({'found' if config_path.exists() else 'MISSING'})")

    main_ok = check_main_py(main_path)

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

    if not main_ok:
        print("\nFix: powershell -ExecutionPolicy Bypass -File update.ps1")
        raise SystemExit(1)


if __name__ == "__main__":
    main()