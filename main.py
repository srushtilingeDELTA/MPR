"""
GSE MPR Report Generator
Excel (May Actuals) -> PowerPoint template with tables and charts
"""

from __future__ import annotations

SCRIPT_VERSION = "2026.07.06-ppt-fill"

import logging
from pathlib import Path

from mpr_data import MprData
from ppt_builder import build_presentation
from report_utils import get_report_period, load_config, parse_bool

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent


def main() -> None:
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
    data.load()

    if data.month_frame(month).empty:
        raise ValueError(
            f"No rows loaded for {report_label}. "
            f"Check sheet '{config['excel']['sheet_name']}' and re-download the workbook."
        )

    output_path = build_presentation(data, config, BASE_DIR)
    print(f"\nDone. Report saved to:\n{output_path}")


if __name__ == "__main__":
    main()