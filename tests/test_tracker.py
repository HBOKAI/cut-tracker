from __future__ import annotations

from datetime import date
from pathlib import Path
import shutil
import subprocess
import sys
import unittest

from openpyxl import load_workbook

from scripts.export_web_data import read_data


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data" / "減脂追蹤.xlsx"
RUNTIME = ROOT / "test_tracker_runtime.xlsx"


class TrackerTests(unittest.TestCase):
    def setUp(self):
        shutil.copy2(SOURCE, RUNTIME)

    def tearDown(self):
        if RUNTIME.exists():
            RUNTIME.unlink()

    def run_log(self, text, *extra):
        return subprocess.run(
            [sys.executable, "scripts/update_daily_log.py", "--workbook", str(RUNTIME), "--input", text, *extra],
            cwd=ROOT,
            capture_output=True,
            text=True,
            encoding="mbcs",
            errors="replace",
        )

    def test_new_day_and_idempotent_repeat(self):
        text = "2026-09-16 訓練日 Calories 2100、Protein 170g、Carbs 200g、Fat 60g"
        first = self.run_log(text)
        self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
        second = self.run_log(text)
        self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
        self.assertIn("沒有需要寫入的變更", second.stdout)

        workbook = load_workbook(RUNTIME, data_only=False)
        worksheet = workbook["每日紀錄"]
        self.assertEqual(worksheet.max_row, 5)
        self.assertEqual(worksheet["A5"].value.date(), date(2026, 9, 16))
        self.assertEqual(worksheet["B5"].value, "TRAINING")
        self.assertEqual(worksheet["F5"].value, 170.0)
        self.assertTrue(str(worksheet["O5"].value).startswith("="))

    def test_same_day_updates_without_duplicate(self):
        first = self.run_log("2026-09-17 休息日 Calories 1800 Protein 150g Carbs 140g Fat 50g")
        self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
        second = self.run_log("2026-09-17 Calories 1900 Protein 165g Carbs 155g Fat 60g")
        self.assertEqual(second.returncode, 0, second.stdout + second.stderr)

        workbook = load_workbook(RUNTIME, data_only=False)
        worksheet = workbook["每日紀錄"]
        matching_rows = [
            row
            for row in range(5, worksheet.max_row + 1)
            if worksheet.cell(row, 1).value.date() == date(2026, 9, 17)
        ]
        self.assertEqual(matching_rows, [6])
        row = matching_rows[0]
        self.assertEqual(worksheet.cell(row, 2).value, "REST")
        self.assertEqual(worksheet.cell(row, 3).value, 1900.0)
        self.assertEqual(worksheet.cell(row, 6).value, 165.0)

    def test_missing_new_day_values_is_rejected(self):
        result = self.run_log("2026-09-18 訓練日")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("請輸入至少一項數值", result.stdout)

    def test_menu_plan_is_available_in_workbook_and_web_data(self):
        workbook = load_workbook(SOURCE, data_only=False)
        self.assertIn("菜單規劃", workbook.sheetnames)
        menu = workbook["菜單規劃"]
        self.assertEqual(menu["A4"].value, "訓練日菜單｜晚上重訓")
        self.assertEqual(menu["A15"].value, "休息日菜單")
        self.assertEqual(menu["A24"].value, "每週安排")

        payload = read_data(SOURCE)
        self.assertEqual(len(payload["menuPlan"]["training"]["rows"]), 6)
        self.assertEqual(len(payload["menuPlan"]["rest"]["rows"]), 3)
        self.assertEqual(payload["menuPlan"]["rest"]["rows"][1]["meal"], "下午茶")
        self.assertIn("樂維根", payload["menuPlan"]["rest"]["rows"][1]["other"])
        self.assertEqual(len(payload["menuPlan"]["weekly"]), 7)
        self.assertEqual(len(payload["menuPlan"]["shopping"]), 7)
        self.assertEqual(len(payload["menuPlan"]["replacements"]), 3)
        self.assertEqual(payload["menuPlan"]["replacements"][0]["item"], "711 鮪魚飯糰")


if __name__ == "__main__":
    unittest.main()
