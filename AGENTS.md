# Agent notes

## Scope

This project is intentionally small. The primary artifact is `data/減脂追蹤.xlsx`, with two sheets only:

- `每日紀錄`: one row per date for Calories, Protein, Carbs, Fat, target minimums, differences, and advice.
- `設定`: training/rest target ranges and a small retained basic-data section.

Do not reintroduce a food database, meal-level input, shopping list, dashboard, or weekly weight log unless the user explicitly asks for it.

## Editing rules

- Use `scripts/update_daily_log.py` for normal daily entries.
- The same date must update the existing row instead of creating a duplicate.
- Create a timestamped workbook backup before every actual write; keep at most 20 backups.
- Keep formulas in target, difference, and advice columns. Ask Excel to recalculate on open.
- Use `openpyxl` for workbook creation and updates.

## Natural-language input

The daily updater accepts Chinese or English labels: `Calories`, `Protein`, `Carbs`, `Fat`, `熱量`, `蛋白質`, `碳水`, `脂肪`; `P / C / F` are also supported. Day type is `TRAINING` or `REST`.

## Verification

After changes, run:

```text
python -m unittest discover -s tests -v
```

Also open or render the workbook to confirm that both sheets are readable and that the advice column is visible.
