# Product Code Formatter Agent Notes

This file is the durable handoff for future Sisyphus/OpenCode sessions. Use it instead of relying on long chat history after compaction.

## Scope And Constraints

- Work in this project root: `C:\Users\shou\Downloads\product_code_flask_app_mst_all_names_v6 (1)`.
- Do not touch `inventory_allocation_app`; it is a separate project.
- The active app is the native PySide desktop app, not the old Flask/browser UI.
- Main entrypoint: `desktop_app.py`.
- Shared business logic: `app.py`.
- Regression tests: `test_process_workbook.py`.
- Build script: `build_exe_auto.bat`.
- Expected EXE: `deploy\ProductCodeFormatter.exe`.
- User prefers PySide: `use Pyside.`
- If a processed output file is already open/locked, the desktop app should clearly tell the user to close the old processed file and retry.

## Current Product Requirements

- Company filtering/skip applies only after the explicit button `Áp dụng lọc công ty`.
- Company table should not show `Trạng thái`.
- Skipped companies appear below separator `Các công ty đã bỏ qua`.
- Product list auto-updates below the selected company; no per-company `Xem hàng hóa` button.
- General `Hàng hóa` is no longer a separate stage.
- `Kiểm tra đơn vị` footer button was removed because it duplicated company filtering.
- Prefix buttons `Áp MST`, `Áp 2 chữ`, `Áp 2 chữ + MST` sit below `Từ thay riêng`.
- `Từ thay riêng` tables should be compact, show about 3 rows, scroll after that, and support deleting rows.
- Output save location must be selectable from the processing flow.
- Cao Thành stage 4 price filtering should match the old web-app price-filter flow.
- Vietmax has two separate profiles: `Vietmax mua vào` and `Vietmax bán ra`.
- Legacy config/profile key `vietmax` maps to `vietmax_mua_vao`.
- `Vietmax mua vào` follows `C:\Users\shou\Documents\Excel Mom\Dóc\VIETMAX\mã vietmax.xlsx` exact/all-word behavior.
- `Vietmax bán ra` code logic: keep the full first word, take 2 characters of remaining normal words, and keep full tokens that contain only uppercase letters and numbers.
- Main processed workbook must add/fill columns `TK vật tư` and `Mã kho`.
- `Mã kho` and `TK vật tư` are coupled as one pair in UI/config/rules.
- User can enter multiple `(Mã kho, TK vật tư)` pairs, pick an optional default pair, and create ordered rules using source column + operator + value + pair.
- Rule operators currently supported: `contains`, `equals`.
- First matching enabled inventory-pair rule wins.
- Multiple rules can reference the same pair to model OR behavior.
- If exactly one inventory pair exists, backend should use that pair for all processed rows even when default checkbox is off.
- Main output default suffix should be `_fdi.xlsx`.
- Companion/second output suffix should be `_nhap_kho.xlsx`, not `_UP_ban_ra.xlsx`.
- Companion `nhap_kho` workbook must also receive `TK vật tư` and `Mã kho`, especially for Vietmax.
- In companion template `mau HD ban ra.xlsx`, `TK vật tư` maps to column `AB`, `Mã kho` maps to column `AF`, and `Mã vật tư` maps to column `AG`.

## Implemented So Far

- Native PySide app is maintained in `desktop_app.py` and build script packages it.
- Profile dropdown includes `vietmax_mua_vao` and `vietmax_ban_ra`; old `vietmax` aliases to `vietmax_mua_vao`.
- `VIETMAX_PROFILES = {"vietmax_mua_vao", "vietmax_ban_ra"}` in `app.py`.
- `Vietmax mua vào` keeps workbook override/all-word behavior.
- `Vietmax bán ra` uses `vietmax_ban_ra_fallback_code(words)`.
- Inventory backend exists in `app.py` with config keys:
  - `inventory_pairs`
  - `use_default_inventory_pair`
  - `default_inventory_pair_id`
  - `inventory_pair_rules`
- Inventory UI exists in `desktop_app.py` on the mapping page.
- Main processed workbook writes `TK vật tư` then `Mã kho` and applies default/rule/single-pair behavior.
- Output chooser exists and mapping/process-page output path fields are synced.
- Long-code warning UI exists for codes over `MAX_CODE_LENGTH = 50`.
- Profile-specific word rules load/save per profile and duplicate detection normalizes accents.
- Cao Thành price filtering stage exists.

## Latest Patch State

The latest suffix/companion/locked-file/reversal patch has been implemented and verified in source tests:

- `app.py`:
  - `resolve_output_path()` default name changed from `{stem}_formatted.xlsx` to `{stem}_fdi.xlsx`.
  - `resolve_output_path()` also normalizes explicitly selected output filenames so they end in `_fdi.xlsx` unless already suffixed.
  - `up_ban_ra_output_path()` changed from `{stem}_UP_ban_ra.xlsx` to `{stem}_nhap_kho.xlsx`.
  - `normalize_inventory_pairs()` now repairs previously saved reversed values such as `ma_kho=152` and `tk_vat_tu=KVT`, writing output as `TK vật tư=152`, `Mã kho=KVT`.
  - Added `processed_inventory_column_indexes(processed_df, output_code_index)` to locate processed `TK vật tư` / `Mã kho` columns by normalized header.
  - `create_up_ban_ra_workbook()` now maps processed `TK vật tư` to template `AB` and processed `Mã kho` to template `AF`.
  - `create_up_ban_ra_workbook()` clears copied template row values before assigning output values, so template sample values like `KHO1` do not leak into generated rows.
- `desktop_app.py`:
  - `ProductCodeFormatterWindow.process_file()` catches `PermissionError` from the main output write and companion `_nhap_kho.xlsx` write, then shows a close-file message with the exact path.
- `test_process_workbook.py`:
  - Imports now include `create_up_ban_ra_workbook`, `resolve_output_path`, and `up_ban_ra_output_path`.
  - Added test for `_fdi` and `_nhap_kho` suffixes.
  - Added test that explicitly selected output filenames are normalized to `_fdi.xlsx`.
  - Added test that reversed saved inventory pairs are normalized before writing the main processed workbook.
  - Added test that generated `nhap_kho` workbook uses `AB2 = TK vật tư` and `AF2 = Mã kho`.

## Remaining Work Checklist

1. Run verification from the project root after any further changes:
   - `.venv\Scripts\python.exe -m py_compile app.py desktop_app.py test_process_workbook.py`
   - `.venv\Scripts\python.exe -m unittest test_process_workbook.py`
2. Rebuild EXE after any further changes:
   - `build_exe_auto.bat`
3. Verify `deploy\ProductCodeFormatter.exe` timestamp/size after rebuild.
4. If tests fail around `nhap_kho` mapping, inspect `create_up_ban_ra_workbook()` and `processed_inventory_column_indexes()` first.

## Last Known Verification Before Latest Patch

- Before the suffix/companion/locked-file patch, these passed:
  - `.venv\Scripts\python.exe -m py_compile app.py desktop_app.py test_process_workbook.py`
  - `.venv\Scripts\python.exe -m unittest test_process_workbook.py`
  - PySide profile smoke
  - Inventory UI payload smoke
- Last rebuilt EXE before latest patch:
  - `deploy\ProductCodeFormatter.exe`
  - timestamp around `2026-06-04 23:40:30`
  - size around `78,642,465` bytes

## Diagnostics Note

- Python LSP may be unavailable unless `basedpyright-langserver` is installed.
- If LSP fails with missing `basedpyright`, rely on compile/tests and mention the LSP limitation in the final verification summary.

## Important Files

- `app.py`: formatter logic, config normalization, profiles, Vietmax rules, inventory pair assignment, output suffixes, companion workbook generation.
- `desktop_app.py`: PySide UI, profile controls, output chooser, inventory pair/rule UI, processing flow and error messages.
- `test_process_workbook.py`: regression tests for workbook processing, Vietmax, inventory pairs, suffixes, and companion workbook mapping.
- `mau HD ban ra.xlsx`: template for companion `nhap_kho` workbook. Relevant columns: `AB = TK vật tư`, `AF = Mã kho`, `AG = Mã vật tư`.
- `build_exe_auto.bat`: packages `desktop_app.py` into `deploy\ProductCodeFormatter.exe`.
