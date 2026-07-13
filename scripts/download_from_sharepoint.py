"""Download MPR Excel from SharePoint into data/.

Methods (pick one):
  sdk        - Office365 REST (usually blocked on Delta — use selenium instead)
  selenium   - Edge/Chrome with manual SSO login (recommended on Delta network)
  local      - Copy from OneDrive-synced folder (no login)

Examples:
  python scripts/download_from_sharepoint.py --method selenium --browser edge
  python scripts/download_from_sharepoint.py --method local
"""

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


def load_sharepoint_config() -> dict:
    config_path = BASE_DIR / "config.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}
    return config.get("sharepoint", {})


def main() -> int:
    parser = argparse.ArgumentParser(description="Download MPR Excel from SharePoint")
    parser.add_argument(
        "--method",
        choices=["sdk", "selenium", "local"],
        default="selenium",
        help="selenium=Edge/Chrome manual login (Delta), local=synced folder",
    )
    parser.add_argument("--save-to", default=None, help="Local path (default: from config.yaml)")
    parser.add_argument("--auth", default="interactive", choices=["interactive", "device"])
    parser.add_argument("--browser", choices=["edge", "chrome"], default=None)
    parser.add_argument("--login-wait", type=int, default=None, help="Seconds to wait for MFA (selenium)")
    parser.add_argument(
        "--edge-path",
        default=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        help="Path to msedge.exe (selenium, Windows)",
    )
    parser.add_argument(
        "--chrome-path",
        default=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        help="Path to chrome.exe (selenium, Windows)",
    )
    args = parser.parse_args()

    sp_cfg = load_sharepoint_config()
    with open(BASE_DIR / "config.yaml", "r", encoding="utf-8") as f:
        excel_cfg = yaml.safe_load(f)["excel"]

    save_to = Path(args.save_to or BASE_DIR / excel_cfg["file_path"])
    file_name = sp_cfg.get("file_name", Path(save_to).name)

    site_url = sp_cfg.get("site_url", "https://deltaairlines.sharepoint.com/sites/DL002488")
    library = sp_cfg.get("library", "GSE MPR Documents")
    folder = sp_cfg.get("folder", "6 - TESTING")
    browser = args.browser or sp_cfg.get("browser", "edge")
    login_wait = args.login_wait if args.login_wait is not None else sp_cfg.get("login_wait_seconds", 180)

    if browser == "edge":
        edge_cfg = sp_cfg.get("edge_path", args.edge_path)
        browser_path = edge_cfg if Path(edge_cfg).exists() else None
        if browser_path is None and Path(args.edge_path).exists():
            browser_path = args.edge_path
        if browser_path is None:
            alt = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            browser_path = alt if Path(alt).exists() else None
    else:
        chrome_cfg = sp_cfg.get("chrome_path", args.chrome_path)
        browser_path = chrome_cfg if Path(chrome_cfg).exists() else None

    try:
        if args.method == "local":
            import sharepoint_excel as spx

            spx.load_excel_local(
                local_path=sp_cfg.get("local_path"),
                file_name=file_name,
                library=library,
                save_to=str(save_to),
            )
            print(f"\nOK: copied/synced file to {save_to}")

        elif args.method == "sdk":
            import sharepoint_excel as spx
            from sharepoint_excel import AuthConfig

            auth = AuthConfig(method=args.auth)
            spx.load_excel_from_sharepoint(
                site_url=site_url,
                library=library,
                folder=folder,
                file_name=file_name,
                auth=auth,
                save_to=str(save_to),
                server_folder_path=sp_cfg.get("server_folder_path"),
            )
            print(f"\nOK: downloaded via Office365 SDK to {save_to}")

        elif args.method == "selenium":
            from sharepoint_selenium import download_via_selenium

            download_via_selenium(
                site_url=site_url,
                library=library,
                folder=folder,
                file_name=file_name,
                save_to=save_to,
                folder_page_url=sp_cfg.get("folder_page_url"),
                server_folder_path=sp_cfg.get("server_folder_path"),
                browser=browser,
                browser_path=browser_path,
                login_wait_seconds=login_wait,
            )
            print(f"\nOK: downloaded via Selenium ({browser}) to {save_to}")

    except Exception as exc:
        logger.error("Download failed: %s", exc)
        print(
            "\nTroubleshooting:\n"
            "  1. Confirm folder_page_url in config.yaml matches your Edge address bar\n"
            "  2. Try --method selenium --browser edge --login-wait 240\n"
            "  3. In Edge: click MPR Actuals and Goals_v2.xlsx → Download manually\n"
            "  4. Try --method local if the library is synced in OneDrive\n"
            "  5. Manual fallback: save file to data/ then run python main.py\n",
            file=sys.stderr,
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
