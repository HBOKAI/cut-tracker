"""Small compatibility helper: update the current-weight snapshot in 設定."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from tracker_utils import DEFAULT_WORKBOOK, backup_before_edit, load_tracker, save_tracker, set_recalculation


def main() -> int:
    parser = argparse.ArgumentParser(description="更新設定頁的目前體重快照")
    parser.add_argument("--weight", type=float, required=True)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--date", help="僅作備註，不建立額外體重歷史列")
    args = parser.parse_args()

    path = args.workbook.resolve()
    try:
        workbook = load_tracker(path)
        worksheet = workbook["設定"]
    except (FileNotFoundError, KeyError) as exc:
        print(f"無法開啟追蹤檔：{exc}")
        return 2
    backup_path = backup_before_edit(path)
    worksheet["B7"] = args.weight
    worksheet["D7"] = f"最近更新 {args.date}" if args.date else "最近更新"
    set_recalculation(workbook)
    save_tracker(workbook, path)
    if backup_path:
        print(f"已備份：{backup_path}")
    print(f"目前體重已更新為 {args.weight:.1f} kg（只保留目前快照）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
