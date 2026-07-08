"""List files in the SharePoint folder (after Edge login)."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config
from sharepoint_live import _authenticate, list_folder_files


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    session, sp_cfg = _authenticate(config)
    site_url = sp_cfg.get("site_url", "https://deltaairlines.sharepoint.com/sites/DL002488")
    folder = sp_cfg.get("server_folder_path", "/sites/DL002488/MPR  Research/GSE MPR Documents/6 - TESTING")

    files = list_folder_files(session, site_url, folder)
    print(f"\nFiles in folder ({len(files)}):\n")
    for row in sorted(files, key=lambda r: (r.get("name") or "").lower()):
        name = row.get("name")
        size = row.get("size")
        print(f"  - {name} ({size:,} bytes)" if size else f"  - {name}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())