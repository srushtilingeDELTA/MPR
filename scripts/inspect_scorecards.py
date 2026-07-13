"""Inspect 2026 - GSE Scorecards.xlsx structure for system scorecard mapping."""

from __future__ import annotations

import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

from report_utils import load_config
from scorecard_data import find_system_scorecard_sheet, load_system_scorecard
from workbook_store import WorkbookStore


class _ScorecardInspector:
    """Minimal data accessor for scorecard-only inspection scripts."""

    def __init__(self, store: WorkbookStore):
        self.store = store

    def sheet_names(self, workbook: str) -> list[str]:
        return self.store.sheet_names(workbook)


def _scorecards_path(config: dict) -> str:
    workbooks = config.get("workbooks", {})
    scorecards = workbooks.get("scorecards", {})
    return scorecards.get("path", "data/2026 - GSE Scorecards.xlsx")


def _try_load_scorecards(store: WorkbookStore) -> bool:
    try:
        store.load(only=["scorecards"])
    except FileNotFoundError:
        return False
    return "scorecards" in store._files


def _load_scorecards_store(config: dict, base_dir: Path) -> WorkbookStore:
    store = WorkbookStore(config, base_dir)
    if _try_load_scorecards(store):
        return store

    sp_cfg = config.get("sharepoint", {})
    if sp_cfg.get("live_read"):
        print("Scorecards file not found locally. Trying SharePoint live read (Edge login may open)...")
        from sharepoint_live import attach_cache_to_config, sync_sharepoint_files

        try:
            result = sync_sharepoint_files(config, base_dir)
            attach_cache_to_config(config, result)
        except Exception as exc:
            print(f"SharePoint sync failed: {exc}")

        store = WorkbookStore(config, base_dir)
        _try_load_scorecards(store)

    return store


def main() -> int:
    config = load_config(base_dir=BASE_DIR)
    store = _load_scorecards_store(config, BASE_DIR)
    if "scorecards" not in store._files:
        scorecards_path = _scorecards_path(config)
        print(f"\nScorecards workbook not available: {scorecards_path}")
        print("\nFix options:")
        print("  1. Save a local copy of the file to:")
        print(f"     {BASE_DIR / scorecards_path}")
        print("  2. Run the full report (SharePoint sync + build):")
        print("     python main.py")
        print("  3. Re-run this script after SharePoint login with live_read enabled")
        return 1

    data = _ScorecardInspector(store)
    names = data.sheet_names("scorecards")
    print(f"Scorecard sheets ({len(names)}):")
    for name in names:
        marker = " <-- likely system scorecard" if name == find_system_scorecard_sheet(names) else ""
        print(f"  - {name}{marker}")

    block = load_system_scorecard(data, block="summary_1")
    if block.empty:
        print("\nNo summary_1 block extracted (Safety, Customer Experience, Operations).")
        return 1

    print(f"\nSummary 1 / slide 3: {len(block)} rows x {len(block.columns)} cols")
    preview = block.head(12).fillna("")
    for i, row in preview.iterrows():
        cells = [str(v)[:18] for v in row.tolist()[:12]]
        print(f"  {i:>3}: {' | '.join(cells)}")

    block2 = load_system_scorecard(data, block="summary_2")
    if block2.empty:
        print("\nNo summary_2 block extracted (Finance, People, footer rows).")
        return 1

    print(f"\nSummary 2 / slide 4: {len(block2)} rows x {len(block2.columns)} cols")
    preview2 = block2.head(12).fillna("")
    for i, row in preview2.iterrows():
        cells = [str(v)[:18] for v in row.tolist()[:12]]
        print(f"  {i:>3}: {' | '.join(cells)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
