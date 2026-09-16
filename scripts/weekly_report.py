"""Print a compact macro summary from 每日紀錄; it does not change the workbook."""

from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path
import statistics
import sys

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from tracker_utils import DEFAULT_WORKBOOK, load_tracker, parse_date


def main() -> int:
    parser = argparse.ArgumentParser(description="列出最近一週的簡單宏量摘要")
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--start", help="週起始日；預設為最近 7 日")
    parser.add_argument("--end", help="週結束日；預設為今天")
    args = parser.parse_args()

    end = parse_date(args.end) if args.end else date.today()
    start = parse_date(args.start) if args.start else end - timedelta(days=6)
    try:
        worksheet = load_tracker(args.workbook.resolve())["每日紀錄"]
    except (FileNotFoundError, KeyError) as exc:
        print(f"無法開啟追蹤檔：{exc}")
        return 2

    rows = []
    for row in range(5, worksheet.max_row + 1):
        value = worksheet[f"A{row}"].value
        if not hasattr(value, "year"):
            continue
        current_date = value.date() if hasattr(value, "date") else value
        if start <= current_date <= end:
            rows.append(row)

    print(f"期間：{start.isoformat()} 至 {end.isoformat()}")
    print(f"已記錄天數：{len(rows)}")
    if not rows:
        print("尚無紀錄。")
        return 0

    columns = {"Calories": "C", "Protein": "F", "Carbs": "I", "Fat": "L"}
    for metric, column in columns.items():
        numbers = [
            float(worksheet[f"{column}{row}"].value)
            for row in rows
            if worksheet[f"{column}{row}"].value is not None
        ]
        if numbers:
            print(f"平均 {metric}: {statistics.mean(numbers):.1f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
