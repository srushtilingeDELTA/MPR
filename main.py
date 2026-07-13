"""
GSE MPR Report Generator
Excel (SharePoint) -> PowerPoint template -> upload to SharePoint
"""

from __future__ import annotations

SCRIPT_VERSION = "2026.07.13-empty-agenda-system-crop"

import logging
from pathlib import Path

from mpr_data import MprData
from ppt_builder import build_presentation
from report_utils import get_report_period, load_config, parse_bool
from sharepoint_live import (
    attach_cache_to_config,
    begin_sharepoint_browser,
    end_sharepoint_browser,
    sync_sharepoint_files,
    upload_file_to_sharepoint,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent


def main() -> None:
    print(f"\nGSE MPR Report Generator  [{SCRIPT_VERSION}]\n")
    config = load_config(base_dir=BASE_DIR)
    report_cfg = config.get("report", {})
    use_previous = parse_bool(report_cfg.get("use_previous_month"), default=True)
    year, month, report_label = get_report_period(config)
    logger.info(
        "Report settings: use_previous_month=%s, Yr_Nb=%s, Mo_Nb=%s (%s)",
        use_previous,
        year,
        month,
        report_label,
    )

    data = MprData(config, BASE_DIR)
    excel_source = data._excel_source()
    logger.info("Excel source: %s", excel_source)

    try:
        if excel_source == "sharepoint":
            logger.info("Opening SharePoint (one Edge window for sync + upload)...")
            begin_sharepoint_browser(config)
            result = sync_sharepoint_files(config, base_dir=BASE_DIR)
            attach_cache_to_config(config, result)
            logger.info(
                "SharePoint sync complete: %s file(s) in memory",
                len(result.files),
            )

        data.load()

        if data.month_frame(month).empty:
            raise ValueError(
                f"No rows loaded for {report_label}. "
                f"Check sheet '{config['excel']['sheet_name']}' and re-download the workbook."
            )

        output_path = build_presentation(data, config, BASE_DIR)
        print(f"\nDone. Report saved to:\n{output_path}")

        sp_cfg = config.get("sharepoint", {})
        if excel_source == "sharepoint" and parse_bool(sp_cfg.get("upload_report"), default=True):
            upload_name = output_path.name
            try:
                server_path = upload_file_to_sharepoint(
                    config,
                    upload_name,
                    output_path.read_bytes(),
                    local_path=output_path,
                )
                print(f"\nUploaded to SharePoint:\n{upload_name}")
                print(f"Server path: {server_path}")
            except Exception as exc:
                logger.error("SharePoint upload failed: %s", exc)
                print(
                    "\nReport was saved locally but SharePoint upload failed.\n"
                    f"Local file: {output_path}\n\n"
                    "Options:\n"
                    "  1. Drag the file into the 6 - TESTING folder in Edge\n"
                    "  2. Run: python scripts\\upload_report_to_sharepoint.py\n"
                )
    finally:
        if excel_source == "sharepoint":
            end_sharepoint_browser(config)


if __name__ == "__main__":
    main()