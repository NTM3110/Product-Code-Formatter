# Product Code Formatter — Agent Instructions

This file is the durable handoff for future AI coding sessions. It describes the project architecture, conventions, build/test processes, and current product constraints. Treat it as the source of truth when starting new work.

## Project Overview

Product Code Formatter is a Windows-first web application that reads Vietnamese Excel invoice workbooks and generates formatted product codes (`Mã VT`) for accounting/ERP import. It groups sellers by `MST` (tax ID), suggests company prefixes, lets users include/skip companies and products, applies profile-specific code-generation rules, and exports a processed Excel workbook plus a companion `nhập kho` workbook.

The active user interface is a React 19 + Vite single-page application that talks to a FastAPI backend. The legacy Angular browser UI and the legacy PySide6 desktop UI have been removed from the working tree; only the React stack and a pywebview desktop wrapper remain.

There are two products in this repository:

1. **Product Code Formatter** (main product) — React 19 + Vite web UI (`react_frontend/`), FastAPI backend (`web_api.py`), and shared Flask backend logic (`app.py`).
2. **Inventory Allocator** (separate product) — Flask backend in `inventory_allocation_app/`, only in scope when explicitly requested. It is integrated into the React UI as stage 12 of the Vietmax workflow.

The repository also contains:

- `frontend/` — Legacy Angular 17 browser UI source. It is no longer the active UI and should not be modified unless the task explicitly targets it.
- `static/` — Last built Angular output, served by the Flask fallback in `app.py`.
- `web_desktop_app.py` — pywebview wrapper around the React/FastAPI stack.

**Do not change the legacy `frontend/` app unless the task explicitly targets the old browser UI.**

## Technology Stack

### Backend / shared logic

- **Python 3.14.4** — script-style codebase, no type annotations required.
- **Flask** (`app.py`) — legacy web API, static serving for the Angular build, shared business logic, and Vietmax processing.
- **FastAPI** (`web_api.py`) — newer API for the React web UI.
- **pandas + openpyxl** — workbook reading/writing and data manipulation.
- **PyInstaller** — single-file Windows EXE packaging for the webview wrapper.
- **Keygen CE (self-hosted)** — license activation and machine fingerprint validation.

### Frontends

| UI | Framework | Entry | Backend | Status |
|---|---|---|---|---|
| React web | React 19 + Vite | `react_frontend/src/vietmax/VietmaxApp.tsx` | `web_api.py` (port 8765) | **Active** |
| Angular web | Angular 17 | `frontend/src/app/app.component.ts` | `app.py` (Flask, port 5000) | Legacy, do not modify |
| Webview desktop | pywebview | `web_desktop_app.py` | `web_api.py` | Current desktop delivery |

### Key Python dependencies

See `requirements.txt` (Product Code Formatter) and `inventory_allocation_app/requirements.txt` (Inventory Allocator).

Product Code Formatter dependencies: `flask`, `flask-cors`, `pandas`, `openpyxl`, `Pillow`, `pyinstaller`, `pystray`, `reportlab`, `fastapi`, `uvicorn`, `python-multipart`, `pywebview`.

Inventory Allocator dependencies: `flask`, `openpyxl`, `pillow`, `pyinstaller`, `pystray`, `PySide6`.

## Repository Layout

```
product_code_flask_app_mst_all_names_v6 (1)/
├── app.py                          # Shared backend: Excel processing, profiles, Vietmax, license helpers
├── web_api.py                      # FastAPI backend for React UI
├── web_desktop_app.py              # pywebview wrapper for React/FastAPI build
├── test_process_workbook.py        # Main unittest regression suite for app.py
├── license_server_admin.py         # Keygen license admin UI (server PC only)
├── create_license_setup_guide_pdf.py  # ReportLab PDF generator
├── ProductCodeFormatterWeb.spec    # PyInstaller spec for webview build
├── requirements.txt
├── mau HD ban ra.xlsx              # Companion template for _nhap_kho.xlsx
├── app_icon.ico
│
├── frontend/                       # Legacy Angular 17 browser UI source
├── react_frontend/                 # React 19 + Vite web UI
├── static/                         # Deployed Angular build served by Flask
├── license_server/                 # Docker Compose scaffold for Keygen CE
│
├── inventory_allocation_app/       # Separate Inventory Allocator product
│   ├── app.py
│   ├── license_client.py
│   ├── test_inventory.py
│   ├── test_license_client.py
│   ├── requirements.txt
│   └── ...
│
├── uploads/                        # Runtime uploaded workbooks
├── outputs/                        # Generated Excel outputs
├── deploy/                         # Shipped EXE artifacts
└── dist/                           # PyInstaller output (copied to deploy/)
```

## Code Organization and Main Modules

### `app.py` — shared business logic (~3,660 lines)

Core areas:

- **Flask app + JSON provider** — `CustomJSONProvider`, legacy API routes, static serving.
- **I/O helpers** — `read_workbook`, `preview_data`, `make_excel_workbook`, `resolve_output_path`, `up_ban_ra_output_path`, `process_zip_stream`.
- **Text/code normalization** — `rm_accents`, `normalize_token`, `code_words`, `excel_col_to_index`, `index_to_excel_col`.
- **Company/product analysis** — `analyze`, `choose_company`, `company_rows`.
- **Code generation** — `make_product_part`, `make_code`, `word_piece`, `phrase_rule_piece`, `apply_word_rules_to_words`, `remove_repeated_phrases`.
- **Profile system** — `PROFILE_LABELS`, `PROFILE_ALIASES`, `profile_key`, `effective_processing_profile`.
- **Vietmax logic** — purchase/sales extraction, matching, conversion formulas, review rows, KHH/152 forcing.
- **Inventory rules** — `normalize_inventory_pairs`, `apply_inventory_pairs`, `processed_inventory_column_indexes`.
- **Price filtering** — Cao Thành price-group/range rules.
- **License / Keygen** — `activate_keygen_license`, `keygen_validate_license`, `keygen_activate_machine`, `license_allows_profile`, `local_machine_fingerprint`.

### `web_api.py` — FastAPI backend for React

Exposes endpoints for the React Vietmax workflow: file upload, mapping, company/prefix, review, matching, processing, export, license activation, config persistence, and Inventory Allocator job management.

### `react_frontend/` — React 19 + Vite web UI

Primary entry: `src/vietmax/VietmaxApp.tsx`. Implements the Vietmax wizard: upload, mapping, company/prefix editing, product review, cross-file matching, processing, and Inventory Allocator stage.

Shared modules:

- `src/api.ts` — HTTP client for FastAPI endpoints.
- `src/types.ts` — shared TypeScript types.
- `src/styles.css` — global styles.
- `src/vietmax/InventoryAllocationStage.tsx` — Inventory Allocator UI.

### `inventory_allocation_app/` — separate product

- `app.py` — allocation engine, ledger/report generation, Flask API.
- `license_client.py` — separate Keygen identity (`inventory-allocator` / `InventoryAllocator`).
- Tests: `test_inventory.py`, `test_license_client.py`.

## Build and Run Commands

### Product Code Formatter (React web)

Run React dev server:

```bat
cd react_frontend
npm run dev
```

Build React bundle:

```bat
cd react_frontend
npm run build
```

Run FastAPI + React backend:

```bat
run_react_app.bat
```

Run webview desktop wrapper:

```bat
run_react_native_app.bat
```

Build webview EXE:

```bat
pyinstaller ProductCodeFormatterWeb.spec --noconfirm
```

After building, copy the result to `deploy/` and rename with a version number:

```bat
Copy-Item dist\ProductCodeFormatterWeb.exe deploy\ProductCodeFormatterWeb_vNN.exe
```

### Product Code Formatter (legacy Angular)

Run Flask + Angular:

```bat
run_python_app.bat
```

The legacy `build_exe_auto.bat` and `ProductCodeFormatter.spec` have been removed from the working tree. Do not reintroduce them.

### Inventory Allocator (only when requested)

Run from source:

```bat
.venv\Scripts\python.exe inventory_allocation_app\app.py
```

Or:

```bat
inventory_allocation_app\run_app.bat
```

### License server admin

```bat
run_license_server_admin.bat
```

## Testing Strategy

### Product Code Formatter

```bat
.venv\Scripts\python.exe -m py_compile app.py web_api.py web_desktop_app.py test_process_workbook.py
.venv\Scripts\python.exe -m unittest test_process_workbook.py
```

`test_process_workbook.py` is the primary regression suite. It covers profile code generation, suffix normalization, inventory pair rules, Vietmax purchase/sales matching, conversion formulas, and license metadata.

### Inventory Allocator

```bat
.venv\Scripts\python.exe -m py_compile inventory_allocation_app\app.py inventory_allocation_app\license_client.py license_server_admin.py
.venv\Scripts\python.exe -m unittest discover inventory_allocation_app
```

## Development Conventions

- **Active UI is React.** The React frontend in `react_frontend/` is the current application; do not reintroduce PySide desktop UI files.
- **Keep `app.py` as the shared backend.** `web_api.py` calls into `app.py`; avoid putting business logic in the UI layer.
- **Configuration lives in `product_code_config.json`.** Use `load_config()` / `save_config()` from `app.py`. Do not edit config files directly.
- **Per-profile config keys** are defined in `empty_profile_config()` in `app.py`. Add new persisted fields there.
- **Inventory pairs** are coupled: `Mã kho` + `TK vật tư`. Rules use `contains` / `equals` operators against source columns.
- **Output suffixes** — main: `_fdi.xlsx`; companion: `_nhap_kho.xlsx`.
- **Vietmax profile** is unified as `vietmax`; legacy keys `vietmax_mua_vao` / `vietmax_ban_ra` remain in code/tests only.
- **Code length warning threshold** is `MAX_CODE_LENGTH = 50`.
- **Do not run `git commit`, `git push`, `git reset`, `git rebase`**, or other git mutations unless explicitly asked.

### PowerShell / Windows command conventions

This project runs on Windows with PowerShell. **Never use `&&` or `||` anywhere in a shell command** — they are not valid statement separators in Windows PowerShell 5.1 and will cause parse errors.

Before every shell tool call, inspect the exact command string. If it contains `&&` or `||`, do not run it. Rewrite it first.

Do not paste generic bash/cmd non-interactive prefixes such as:

```powershell
set CI="true" && set GIT_TERMINAL_PROMPT="0" && git ...
```

Use PowerShell environment assignment syntax instead, and separate statements with semicolons:

```powershell
$env:CI="true"; $env:GIT_TERMINAL_PROMPT="0"; git ...
```

Use one of these approaches instead:

1. **Semicolon `;`** — runs commands sequentially regardless of success:
   ```powershell
   cmd1; cmd2; cmd3
   ```

2. **Separate tool calls** — preferred for independent commands.

3. **Conditional with `if ($?)`** — for dependent commands:
   ```powershell
   cmd1; if ($?) { cmd2 }
   ```

4. **Pipeline with `|`** — when appropriate:
   ```powershell
   Get-Content file.txt | Select-String pattern
   ```

### Preventing command loop bugs

When a shell command fails (especially PowerShell syntax errors):

1. **Stop immediately** — do not retry the same command with the same syntax.
2. **Analyze the error** — read the error message and identify the actual problem.
3. **Fix the root cause** — change the approach, not just retry.
4. **Use correct PowerShell syntax** — never use `&&` or `||`.

Common loop patterns to avoid:
- Repeated searches with the same query returning no results.
- Retrying failed commands without changing the syntax.
- Using `&&` in PowerShell after being told it does not work.
- Continuing to search when the file does not exist or the pattern is wrong.

## Security Considerations

- The app requires Keygen license activation before normal use. License state is stored in `product_code_config.json`.
- Activation is machine-bound: `local_machine_fingerprint()` uses `COMPUTERNAME` + MAC address. Moving the config to another PC requires reactivation.
- Keygen admin tokens and secrets must **never** be embedded in the desktop app. Only the license key and activation result are stored locally.
- HTTP Keygen URLs are allowed only for localhost, loopback, private/LAN IPs, `.local` names, and single-label LAN hostnames. Public HTTP hosts are rejected.
- The client sends `X-Forwarded-Proto: https` for allowed HTTP hosts because the Keygen CE container redirects to HTTPS.
- License metadata can restrict allowed profiles via keys: `allowed_profiles`, `profiles`, `company_profiles`, `allowed_companies`, `companies`.
- Vietmax conversion formulas are parsed only in simple safe ratio form (e.g. `1 ram = 500 tờ`); arbitrary expression evaluation is not allowed.
- Do not commit `.env` files, license keys, or private certificates. These are already ignored in `.gitignore`.

## License and Activation

The license server is a self-hosted `keygen-sh/keygen-api` instance on the same LAN, typically hosted from Windows with Docker Desktop.

- Server scaffold: `license_server/docker-compose.yml` + `.env`.
- Setup notes: `license_server/README.md`.
- Admin UI: `license_server_admin.py` (run only on the server PC).
- Client activation flow:
  1. React UI collects server URL, account id/slug, and license key.
  2. `activate_keygen_license()` validates the key; if the machine is not yet activated, it creates the machine relationship and revalidates.
  3. On success, `activated=true` and the machine fingerprint are saved.
  4. Later launches skip the server if the saved fingerprint matches the current PC.
- License metadata restricts which company profiles appear in the profile dropdown.

## Current Product Requirements and Constraints

The following requirements are active and must be preserved:

- Company filtering/skip applies only after the explicit button `Áp dụng lọc công ty`.
- Company table does not show `Trạng thái`.
- Skipped companies appear below the separator `Các công ty đã bỏ qua`.
- Product list auto-updates below the selected company; no per-company `Xem hàng hóa` button.
- Prefix buttons `Áp MST`, `Áp 2 chữ`, `Áp 2 chữ + MST` sit below `Từ thay riêng`.
- `Từ thay riêng` tables are compact (~3 rows), scroll after that, and support deleting rows.
- Output save location is selectable from the processing flow.
- Cao Thành price filtering is stage 6 in the web UI, after company/prefix and post-company `Review Mã VT`.
- Unified Vietmax has two phases: `purchase` (`vietmax_mua_vao` logic, forces `KHH/152`), then `sales` (`vietmax_ban_ra` logic, buyer-side labels, defaults `I/J/K/M/O`, output `L`, invoice status `AJ`, no price column).
- Stage `2. Khớp HD mua vào / bán ra` is visible for all profiles after selecting the main workbook. Vietmax shows 3 tabs (`HD mua vào` review, `HD bán ra` review, `Khớp mua/bán`); non-Vietmax shows only the match tab.
- Exact-duplicate product names are hidden from internal review tabs; only near-similar names are shown.
- Internal-review rows that differ only by case/space are auto-confirmed and grouped under `Chỉ khác nhau cách viết`; rows that differ only by dimensions are skipped entirely.
- Cross-file match table shows `Số HD`, `Ngày có hàng bán ra`, `ĐVT bán`, `ĐVT mua`, mismatch warnings, and editable conversion formulas.
- Confirmed mismatched matches support per-row conversion modes: none, quantity+unit, quantity only, unit only.
- Unchecked/skipped rows must not receive purchase-code override, conversion, or forced `KHH/152`.
- Main processed workbook must add/fill `TK vật tư` and `Mã kho`.
- Companion `nhap_kho` workbook maps `TK vật tư` → `AB`, `Mã kho` → `AF`, `Mã vật tư` → `AG` in template `mau HD ban ra.xlsx`.
- Company prefix suggestion removes Vietnamese province/city names and `Việt Nam` / `Viet Nam` before taking initials.
- If a processed output file is open/locked, the web app tells the user to close it and retry.

## Important Files

| File | Purpose |
|---|---|
| `app.py` | Formatter backend: profiles, code generation, Vietmax rules, inventory pairs, output suffixes, license helpers |
| `web_api.py` | FastAPI backend for React Vietmax UI |
| `web_desktop_app.py` | pywebview wrapper around React/FastAPI |
| `test_process_workbook.py` | Main regression tests for backend |
| `ProductCodeFormatterWeb.spec` | PyInstaller spec for current webview build |
| `mau HD ban ra.xlsx` | Companion `nhap_kho` template |
| `app_icon.ico` | Icon used by webview wrapper |
| `license_server_admin.py` | Keygen license creation UI |
| `create_license_setup_guide_pdf.py` | PDF setup guide generator |
| `inventory_allocation_app/app.py` | Inventory Allocator backend |
| `react_frontend/src/vietmax/VietmaxApp.tsx` | React Vietmax wizard |
| `react_frontend/src/styles.css` | Global styles |

## Verification Checklist

After any Product Code Formatter backend change:

```bat
.venv\Scripts\python.exe -m py_compile app.py web_api.py web_desktop_app.py test_process_workbook.py
.venv\Scripts\python.exe -m unittest test_process_workbook.py
```

After any React change:

```bat
cd react_frontend
npm run build
```

After any Inventory Allocator change:

```bat
.venv\Scripts\python.exe -m py_compile inventory_allocation_app\app.py inventory_allocation_app\license_client.py license_server_admin.py
.venv\Scripts\python.exe -m unittest discover inventory_allocation_app
```

After building the webview EXE, smoke test it with a pinned port:

```bat
$env:PRODUCT_CODE_FORMATTER_PORT="5000"
.\deploy\ProductCodeFormatterWeb.exe
# In another terminal: curl http://127.0.0.1:5000/api/health
```

## Context Compaction Guide

When context is overloaded, summarize only the minimum needed to continue work.

Use this format:

```text
Goal:
- one sentence

Changed Files:
- path + one short reason each

Current Behavior:
- 3-6 bullets only

Open Work:
- exact next tasks only

Verification:
- latest build/test status only

Known Constraints:
- only active constraints still relevant
```

Rules for compact handoff:
- do not paste long config examples unless the bug is about config shape
- do not repeat full user history
- do not list unchanged files
- prefer symbol names over code excerpts
- mention only the latest accepted behavior, not every prior attempt
- if a UI area is affected, name the modal/section instead of describing the whole screen
- if build/test already passed, record only the latest result

Good compact summary example:

```text
Goal:
- Fix Cao Thành price-group modal behavior and rebuild EXE.

Changed Files:
- frontend/src/app/app.component.ts - price grouping, skipped-item persistence
- frontend/src/app/app.component.html - price modal controls
- frontend/src/app/app.component.spec.ts - regression tests
- app.py - config persistence semantics

Current Behavior:
- price modal groups by final customized Mã VT
- skipped products are stored as skipped, not selected
- product modal shows code length and >50 warning

Open Work:
- rebuild deploy EXE after frontend change

Verification:
- npm run build passed
- npm test passed: 22 SUCCESS

Known Constraints:
- CSS budget warning exists but does not block build
```

## Compaction Priority

If you must cut more context, keep information in this order:

1. current goal
2. changed files
3. exact unfinished work
4. verification state
5. only critical constraints

Drop everything else first.

## AI Agent Configuration

The following settings are permanent defaults for AI agent behavior in this project:

- **Search timeout**: Stop searching after 3-5 attempts or 30 seconds, whichever comes first. Do not commit to long-running searches without user confirmation.
- **Context window**: When context exceeds 70% capacity, trigger compaction immediately using the Context Compaction Guide above.
- **Subagent usage**: Only use subagents for parallel exploration or complex multi-step tasks. For simple lookups, use direct tools (Read, Grep, AST-grep) instead.
- **Search strategy**: Try direct file reads first, then Grep/AST-grep for patterns. Only use subagent search if the codebase is unfamiliar or the pattern spans multiple files.
- **Error handling**: After 2 failed attempts on the same approach, switch to a different strategy. Do not retry the same failing command more than twice.
- **Build verification**: Always run `npm run build` after React changes and `py_compile` after Python changes before declaring work complete.
