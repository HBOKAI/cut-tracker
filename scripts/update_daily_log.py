"""Add or update one day's four macro totals."""

from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import re
import sys

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from tracker_utils import (
    DEFAULT_WORKBOOK,
    backup_before_edit,
    date_from_text,
    find_daily_row,
    infer_day_type,
    load_tracker,
    parse_date,
    save_tracker,
    set_recalculation,
    target_min,
    update_daily_table_ref,
    write_daily_formulas,
)


METRIC_PATTERNS = {
    "Calories": [r"(?:Calories?|熱量|kcal)\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)"],
    "Protein": [
        r"(?:Protein|蛋白質|蛋白)\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
        r"\bP\b\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
    ],
    "Carbs": [
        r"(?:Carbs?|碳水(?:化合物)?)\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
        r"\bC\b\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
    ],
    "Fat": [
        r"(?:Fat|脂肪)\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
        r"\bF\b\s*[:：=]?\s*([0-9]+(?:[.,][0-9]+)?)",
    ],
}


def extract_metric(text: str, metric: str, override: float | None) -> float | None:
    if override is not None:
        return float(override)
    for pattern in METRIC_PATTERNS[metric]:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return float(match.group(1).replace(",", "."))
    return None


def advice(day_type: str | None, values: dict[str, float | None], settings) -> str:
    if not day_type:
        return "請先填 Day_Type（TRAINING 或 REST）。"
    if any(values.get(metric) is None for metric in ("Calories", "Protein", "Carbs", "Fat")):
        return "請補齊 Calories、Protein、Carbs、Fat。"
    diffs = {
        metric: target_min(settings, day_type, metric) - float(values[metric])
        for metric in ("Calories", "Protein", "Carbs", "Fat")
    }
    if diffs["Protein"] > 0:
        return "優先補蛋白質。"
    if diffs["Carbs"] > 0:
        return "蛋白質已足夠，優先補碳水，例如白飯或香蕉。"
    if diffs["Fat"] > 0:
        return "優先補脂肪。"
    if diffs["Calories"] > 0:
        return "三大營養素已達最低，熱量尚差；可補少量主食。"
    return "已達每日最低目標；蛋白質足夠，不需補充乳清。"


def format_number(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{float(value):.1f}".rstrip("0").rstrip(".")


def main() -> int:
    parser = argparse.ArgumentParser(description="用自然語言記錄每日 Calories / Protein / Carbs / Fat")
    parser.add_argument(
        "--input",
        "-i",
        required=True,
        help="例如：今天訓練日，Calories 2100、Protein 170g、Carbs 220g、Fat 60g",
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--date", help="覆蓋文字中的日期；可用 YYYY-MM-DD、今天、昨天")
    parser.add_argument("--day-type", choices=["TRAINING", "REST"], help="覆蓋文字中的 Day_Type")
    parser.add_argument("--calories", type=float)
    parser.add_argument("--protein", type=float)
    parser.add_argument("--carbs", type=float)
    parser.add_argument("--fat", type=float)
    args = parser.parse_args()

    text = args.input
    try:
        target_date = parse_date(args.date) if args.date else date_from_text(text, default=date.today())
    except ValueError as exc:
        print(str(exc))
        return 2

    day_type = args.day_type or infer_day_type(text)
    values = {
        "Calories": extract_metric(text, "Calories", args.calories),
        "Protein": extract_metric(text, "Protein", args.protein),
        "Carbs": extract_metric(text, "Carbs", args.carbs),
        "Fat": extract_metric(text, "Fat", args.fat),
    }

    workbook_path = args.workbook.resolve()
    try:
        workbook = load_tracker(workbook_path)
        worksheet = workbook["每日紀錄"]
        settings = workbook["設定"]
    except (FileNotFoundError, KeyError) as exc:
        print(f"無法開啟簡化版追蹤檔：{exc}")
        return 2

    row = find_daily_row(worksheet, target_date)
    if row is None and all(value is None for value in values.values()):
        print("請輸入至少一項數值；建議格式：今天訓練日，Calories 2100、Protein 170g、Carbs 220g、Fat 60g")
        return 2

    is_new_row = row is None
    if is_new_row:
        first_data_row_is_empty = worksheet.max_row <= 5 and all(
            worksheet.cell(5, column).value is None for column in range(1, 17)
        )
        row = 5 if first_data_row_is_empty else max(4, worksheet.max_row) + 1
        current = {"Calories": None, "Protein": None, "Carbs": None, "Fat": None}
        current_day_type = None
    else:
        current = {
            "Calories": worksheet[f"C{row}"].value,
            "Protein": worksheet[f"F{row}"].value,
            "Carbs": worksheet[f"I{row}"].value,
            "Fat": worksheet[f"L{row}"].value,
        }
        current_day_type = worksheet[f"B{row}"].value

    new_values = {metric: (values[metric] if values[metric] is not None else current[metric]) for metric in values}
    new_day_type = day_type or current_day_type or ""
    changed = is_new_row or new_day_type != (current_day_type or "")
    changed = changed or any(new_values[metric] != current[metric] for metric in values)
    if not changed:
        print(f"{target_date.isoformat()} 沒有需要寫入的變更。")
        return 0

    backup_path = backup_before_edit(workbook_path)
    worksheet[f"A{row}"] = target_date
    worksheet[f"B{row}"] = new_day_type
    worksheet[f"C{row}"] = new_values["Calories"]
    worksheet[f"F{row}"] = new_values["Protein"]
    worksheet[f"I{row}"] = new_values["Carbs"]
    worksheet[f"L{row}"] = new_values["Fat"]
    write_daily_formulas(worksheet, row)
    update_daily_table_ref(worksheet)
    set_recalculation(workbook)
    save_tracker(workbook, workbook_path)

    actual_advice = advice(new_day_type or None, new_values, settings)
    if backup_path:
        print(f"已備份：{backup_path}")
    print(f"已更新 {target_date.isoformat()}（{new_day_type or '未填 Day_Type'}）")
    print("  Calories: " + format_number(new_values["Calories"]) + " / " + (format_number(target_min(settings, new_day_type, "Calories")) if new_day_type else "-"))
    print("  Protein:  " + format_number(new_values["Protein"]) + " / " + (format_number(target_min(settings, new_day_type, "Protein")) if new_day_type else "-"))
    print("  Carbs:    " + format_number(new_values["Carbs"]) + " / " + (format_number(target_min(settings, new_day_type, "Carbs")) if new_day_type else "-"))
    print("  Fat:      " + format_number(new_values["Fat"]) + " / " + (format_number(target_min(settings, new_day_type, "Fat")) if new_day_type else "-"))
    print("  建議：" + actual_advice)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
