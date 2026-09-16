# 減脂飲食追蹤系統

這是簡化版的長期追蹤檔，日常只需要記錄四個數字：Calories、Protein、Carbs、Fat。

檔案位置：`data/減脂追蹤.xlsx`

## 怎麼用

最簡單的方式是在終端機輸入：

```text
python scripts/update_daily_log.py --input "今天訓練日，Calories 2100、Protein 170g、Carbs 220g、Fat 60g"
```

程式會在 `每日紀錄` 每天保留一列；同一天再次輸入會更新原列，不會重複新增。輸入可用中文或英文欄位名稱，也可用 `P / C / F`。

工作簿只有兩頁：

- `每日紀錄`：實際攝取、目標最低值、差額與一句建議。
- `設定`：訓練日與休息日的目標範圍，以及目前保留的基本資料。

差額定義為「目標最低值 − 實際」。正數表示還差多少；Protein 已達最低值時，建議不會叫你補乳清。

## 其他指令

更新設定頁的目前體重快照：

```text
python scripts/update_weight.py --weight 76.2 --date 2026-09-17
```

列出最近 7 日的平均宏量：

```text
python scripts/weekly_report.py
```

每次實際寫入前會在 `backups/` 保留備份，最多保留 20 份。測試指令：

```text
python -m unittest discover -s tests -v
```

## GitHub Pages

`docs/` 是可以直接發布的靜態儀表板，會顯示四項攝取、目標差額、建議與最近 7 次 Calories 趨勢。網站只匯出宏量資料，不匯出設定頁的身高、體重與體脂率。

請將這個資料夾推送到 GitHub 儲存庫，然後在儲存庫的 **Settings → Pages** 將來源設為 **GitHub Actions**。`.github/workflows/pages.yml` 會自動讀取 `data/減脂追蹤.xlsx`、產生 `docs/data.json` 並發布頁面。

本機若要先更新網頁資料，可執行：

```text
python scripts/export_web_data.py
```

如果儲存庫是公開的，GitHub Pages 上的飲食紀錄也會公開；建議使用私有儲存庫，或確認你接受這些資料被看見。
