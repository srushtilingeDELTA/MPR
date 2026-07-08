"""Fetch all configured SharePoint files in one Edge login."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config
from sharepoint_live import attach_cache_to_config, sync_sharepoint_files


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    result = sync_sharepoint_files(config, base_dir=BASE_DIR)
    attach_cache_to_config(config, result)

    print(f"\nLoaded {len(result.files)} file(s) from SharePoint:")
    for dest, data in result.files.items():
        print(f"  - {dest} ({len(data):,} bytes)")

    if result.errors:
        print(f"\nSkipped {len(result.errors)} optional/error file(s):")
        for name, err in result.errors.items():
            print(f"  - {name}: {err}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())