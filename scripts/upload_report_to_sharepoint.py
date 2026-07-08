"""Upload the latest generated report to SharePoint (one Edge window)."""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import get_report_period, load_config
from sharepoint_live import (
    begin_sharepoint_browser,
    end_sharepoint_browser,
    upload_file_to_sharepoint,
)


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    year, month, _ = get_report_period(config)
    month_name = date(year, month, 1).strftime("%B")
    output_name = config["powerpoint"].get("output_name", "GSE MPR - {month_name} {year}.pptx").format(
        year=year,
        month=month,
        month_name=month_name,
    )
    output_path = BASE_DIR / config["powerpoint"]["output_dir"] / output_name
    if not output_path.exists():
        print(f"Report not found: {output_path}\nRun python main.py first.")
        return 1

    try:
        begin_sharepoint_browser(config)
        server_path = upload_file_to_sharepoint(
            config,
            output_path.name,
            output_path.read_bytes(),
            local_path=output_path,
        )
        print(f"Uploaded: {output_path.name}")
        print(f"Server path: {server_path}")
        return 0
    finally:
        end_sharepoint_browser(config)


if __name__ == "__main__":
    raise SystemExit(main())