"""Export the macro-only portion of the workbook for GitHub Pages."""

from __future__ import annotations

import argparse
from datetime import date, datetime
import json
from pathlib import Path
import sys

from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
DEFAULT_WORKBOOK = ROOT / "data" / "減脂追蹤.xlsx"
DEFAULT_OUTPUT = ROOT / "docs" / "data.json"
TARGET_ROWS = {"Calories": 12, "Protein": 13, "Carbs": 14, "Fat": 15}
MENU_HEADERS = ["time", "meal", "chicken", "rice", "riceCup", "other", "greens", "note"]


def as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def number(value) -> float | None:
    if value is None or value == "":
        return None
    return float(value)


def advice(actual: dict[str, float | None], target: dict[str, float]) -> str:
    if any(actual.get(metric) is None for metric in TARGET_ROWS):
        return "請補齊 Calories、Protein、Carbs、Fat。"
    diff = {metric: target[metric] - actual[metric] for metric in TARGET_ROWS}
    if diff["Protein"] > 0:
        return "優先補蛋白質。"
    if diff["Carbs"] > 0:
        return "蛋白質已足夠，優先補碳水，例如白飯或香蕉。"
    if diff["Fat"] > 0:
        return "優先補脂肪。"
    if diff["Calories"] > 0:
        return "三大營養素已達最低，熱量尚差；可補少量主食。"
    return "已達每日最低目標；蛋白質足夠，不需補充乳清。"


def display_text(value) -> str:
    return "" if value is None else str(value)


def read_menu_plan(workbook) -> dict | None:
    """Read the compact menu reference sheet for the static website."""

    if "菜單規劃" not in workbook.sheetnames:
        return None
    menu = workbook["菜單規劃"]

    def rows(start: int, end: int) -> list[dict[str, str]]:
        return [
            {key: display_text(menu.cell(row, column).value) for column, key in enumerate(MENU_HEADERS, 1)}
            for row in range(start, end + 1)
        ]

    def total(row: int) -> dict[str, str]:
        return {
            "chicken": display_text(menu.cell(row, 3).value),
            "rice": display_text(menu.cell(row, 4).value),
            "riceCup": display_text(menu.cell(row, 5).value),
            "other": display_text(menu.cell(row, 6).value),
            "greens": display_text(menu.cell(row, 7).value),
        }

    weekly = []
    for row in range(26, 33):
        weekly.append({
            "day": display_text(menu.cell(row, 1).value),
            "type": display_text(menu.cell(row, 2).value),
            "chicken": display_text(menu.cell(row, 3).value),
            "rice": display_text(menu.cell(row, 4).value),
            "mackerel": display_text(menu.cell(row, 5).value),
            "eggs": display_text(menu.cell(row, 6).value),
        })

    shopping = []
    for row in range(37, 44):
        shopping.append({
            "item": display_text(menu.cell(row, 1).value),
            "quantity": display_text(menu.cell(row, 2).value),
        })

    return {
        "training": {
            "title": display_text(menu.cell(4, 1).value),
            "rows": rows(6, 11),
            "total": total(12),
        },
        "rest": {
            "title": display_text(menu.cell(15, 1).value),
            "rows": rows(17, 20),
            "total": total(21),
        },
        "weekly": weekly,
        "shopping": shopping,
    }


def read_data(workbook_path: Path) -> dict:
    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        settings = workbook["設定"]
        daily = workbook["每日紀錄"]
    except KeyError as exc:
        raise ValueError("工作簿需要包含「每日紀錄」與「設定」兩頁。") from exc

    target_ranges = {"TRAINING": {}, "REST": {}}
    for metric, row in TARGET_ROWS.items():
        target_ranges["TRAINING"][metric] = {"min": number(settings.cell(row, 2).value), "max": number(settings.cell(row, 3).value)}
        target_ranges["REST"][metric] = {"min": number(settings.cell(row, 4).value), "max": number(settings.cell(row, 5).value)}

    days = []
    for row in range(5, daily.max_row + 1):
        current_date = as_date(daily.cell(row, 1).value)
        if current_date is None:
            continue
        day_type = daily.cell(row, 2).value or ""
        actual = {
            "Calories": number(daily.cell(row, 3).value),
            "Protein": number(daily.cell(row, 6).value),
            "Carbs": number(daily.cell(row, 9).value),
            "Fat": number(daily.cell(row, 12).value),
        }
        target = {
            metric: target_ranges[day_type][metric]["min"]
            for metric in TARGET_ROWS
            if day_type in target_ranges
        }
        diff = {
            metric: (target[metric] - actual[metric]) if metric in target and actual[metric] is not None else None
            for metric in TARGET_ROWS
        }
        days.append({
            "date": current_date.isoformat(),
            "dayType": day_type,
            "actual": actual,
            "target": target,
            "diff": diff,
            "advice": advice(actual, target) if len(target) == len(TARGET_ROWS) else "請先填 Day_Type（TRAINING 或 REST）。",
        })

    days.sort(key=lambda item: item["date"], reverse=True)
    return {
        "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "targetRanges": target_ranges,
        "menuPlan": read_menu_plan(workbook),
        "days": days,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="匯出 GitHub Pages 使用的宏量資料")
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    try:
        payload = read_data(args.workbook.resolve())
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc))
        return 2
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已匯出網頁資料：{output}")
    print(f"包含每日紀錄：{len(payload['days'])} 天")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
