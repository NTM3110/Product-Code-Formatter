# Fast Change Guide

Use this guide to keep small fixes small.

## Before Editing

1. Identify the exact stage/profile.
2. Search by label or endpoint with `rg`.
3. Edit the smallest layer that owns the behavior.
4. Do not move large files or refactor unrelated code during a bug fix.

## Stage Map

Vietmax stages:

- `1` upload purchase
- `2` choose purchase columns
- `3` purchase company/prefix/product selection
- `4` purchase Review Ma VT
- `5` create purchase file/cache
- `6` upload sales
- `7` choose sales columns
- `8` match purchase/sales
- `9` sales company/prefix/product selection
- `10` sales Review Ma VT
- `11` create sales file/cache
- `12` inventory allocation
- `13` report view
- `14` export allocation file

## Quick Validation

After React-only changes:

```powershell
Set-Location -LiteralPath "E:\Excel Mom\product_code_flask_app_mst_all_names_v6 (1)\react_frontend"
npm run build
```

After Python-only changes:

```powershell
Set-Location -LiteralPath "E:\Excel Mom\product_code_flask_app_mst_all_names_v6 (1)"
$env:PYTHONDONTWRITEBYTECODE='1'
.\.venv\Scripts\python.exe -B -c "from pathlib import Path; files=['app.py','web_api.py','web_desktop_app.py','test_process_workbook.py']; [compile(Path(f).read_text(encoding='utf-8-sig'), f, 'exec') for f in files]; print('syntax ok')"
```

Only run the full unittest suite when the change touches workbook logic or when explicitly requested. It can fail if Excel/test files are open.

## Rebuild Desktop EXE

```powershell
Set-Location -LiteralPath "E:\Excel Mom\product_code_flask_app_mst_all_names_v6 (1)"
.\.venv\Scripts\python.exe -m PyInstaller ProductCodeFormatterWeb.spec --noconfirm
Copy-Item -LiteralPath .\dist\ProductCodeFormatterWeb.exe -Destination .\deploy\ProductCodeFormatterWeb.exe -Force
```

If copy fails because the exe is running:

```powershell
Get-Process | Where-Object { $_.ProcessName -like '*ProductCodeFormatter*' } | Select-Object Id,ProcessName,Path
Stop-Process -Id <id> -Force
```

## Prefix Config Checklist

When changing company prefix behavior, keep purchase and sales separate:

- Purchase state: `purchasePrefixStrategy`, `purchasePrefixStrategyValues`
- Sales state: `salesPrefixStrategy`, `salesPrefixStrategyValues`
- Backend profile storage:
  - purchase: `vietmax_mua_vao`
  - sales: `vietmax_ban_ra`

Each prefix option must have its own manual values:

- `last_2_words`
- `last_3_mst`
- `2_words_mst`

Save/load fields:

- `prefix_strategy`
- `prefix_mst_digits`
- `prefix_strategy_values`

## Avoid These Time Sinks

- Do not rebuild the EXE unless the user needs the desktop artifact.
- Do not run full tests for UI-only text/layout changes.
- Do not edit `frontend/` unless the old Angular UI is explicitly requested.
- Do not manually clean generated folders during feature work.
- Do not use broad git operations in this repo because the worktree often has unrelated changes.

