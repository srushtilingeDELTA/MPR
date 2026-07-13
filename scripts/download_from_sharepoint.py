"""Download MPR Excel from SharePoint into data/ (Selenium or live sync)."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import yaml

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)


def main() -> int:
    parser = argparse.ArgumentParser(description="Download MPR Excel from SharePoint")
    parser.add_argument(
        "--method",
        choices=["selenium", "live"],
        default="live",
        help="live = Edge + REST (recommended); selenium = browser UI download",
    )
    parser.add_argument("--save-to", default=None)
    parser.add_argument("--browser", choices=["edge", "chrome"], default=None)
    parser.add_argument("--login-wait", type=int, default=None)
    args = parser.parse_args()

    with open(BASE_DIR / "config.yaml", "r", encoding="utf-8") as handle:
        config = yaml.safe_load(handle) or {}
    sp_cfg = config.get("sharepoint", {})
    excel_cfg = config.get("excel", {})

    save_to = Path(args.save_to or BASE_DIR / excel_cfg["file_path"])
    save_to.parent.mkdir(parents=True, exist_ok=True)

    try:
        if args.method == "live":
            from sharepoint_live import (
                attach_cache_to_config,
                begin_sharepoint_browser,
                end_sharepoint_browser,
                sync_sharepoint_files,
            )

            begin_sharepoint_browser(config)
            try:
                result = sync_sharepoint_files(config, base_dir=BASE_DIR)
                attach_cache_to_config(config, result)
                # Prefer writing the primary actuals workbook to disk.
                dest_key = excel_cfg.get("file_path", "data/MPR Actuals and Goals_v2.xlsx")
                payload = result.files.get(dest_key) or next(iter(result.files.values()), None)
                if not payload:
                    raise RuntimeError("Live sync returned no files.")
                save_to.write_bytes(payload)
                print(f"\nOK: saved via live sync to {save_to}")
            finally:
                end_sharepoint_browser(config)
            return 0

        from sharepoint_selenium import download_via_selenium

        browser = args.browser or sp_cfg.get("browser", "edge")
        login_wait = (
            args.login_wait
            if args.login_wait is not None
            else sp_cfg.get("login_wait_seconds", 180)
        )
        browser_path = sp_cfg.get("edge_path") if browser == "edge" else sp_cfg.get("chrome_path")
        if browser_path and not Path(browser_path).exists():
            browser_path = None

        download_via_selenium(
            site_url=sp_cfg.get("site_url", "https://deltaairlines.sharepoint.com/sites/DL002488"),
            library=sp_cfg.get("library", "GSE MPR Documents"),
            folder=sp_cfg.get("folder", "6 - TESTING"),
            file_name=sp_cfg.get("file_name", save_to.name),
            save_to=save_to,
            folder_page_url=sp_cfg.get("folder_page_url"),
            server_folder_path=sp_cfg.get("server_folder_path"),
            browser=browser,
            browser_path=browser_path,
            login_wait_seconds=login_wait,
        )
        print(f"\nOK: downloaded via Selenium ({browser}) to {save_to}")
        return 0
    except Exception as exc:
        logger.error("Download failed: %s", exc)
        print(
            "\nTroubleshooting:\n"
            "  1. Confirm folder_page_url in config.yaml\n"
            "  2. Try --login-wait 240\n"
            "  3. Prefer: python main.py (live SharePoint path)\n",
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
