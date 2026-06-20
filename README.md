# Product Code Formatter - MST All Names V6

## Overview

Windows desktop app for formatting `Mã VT` from Excel invoice data. The active UI is React + Vite, served by a FastAPI backend and wrapped in pywebview for the desktop `.exe`.

Main capabilities:
- require LAN Keygen license activation before normal use
- group sellers by `MST`
- suggest and validate company prefixes
- let users include or skip companies and products
- generate product codes by profile (`Sơn Phương`, `Cao Thành`, `Quang Thịnh`, `Vietmax`)
- persist config, rules, overrides, price-group settings, and skipped items
- export a processed Excel workbook

## Current Notes

- Active UI: `react_frontend/`
- Active backend: `web_api.py`
- Shared Excel/business logic: `app.py`
- Desktop wrapper: `web_desktop_app.py`
- Future edit map: `docs/PROJECT_MAP.md`
- Fast fix checklist: `docs/FAST_CHANGE_GUIDE.md`
- The company check page shows all company names found under the same `MST`.
- The app still groups by seller `MST`.
- Prefix suggestion uses the most common company name for that `MST`.
- Users can tick or untick which company groups and products to process.
- Rows where quantity is empty, invalid, or `0` are skipped.

## Build

1. Extract the project ZIP.
2. Build the React bundle: `cd react_frontend; npm run build`.
3. Build the desktop wrapper: `.\.venv\Scripts\python.exe -m PyInstaller ProductCodeFormatterWeb.spec --noconfirm`.
4. Copy `dist\ProductCodeFormatterWeb.exe` to `deploy\ProductCodeFormatterWeb.exe`.

## License Activation

The desktop app shows a license activation dialog before the main window opens. Use a self-hosted `keygen-sh/keygen-api` server on the same LAN and enter:

- License server URL, such as `http://license-server.local:3000` for private LAN use or `https://license-server.local` when using an HTTPS reverse proxy
- Keygen account id or slug
- License key

For local setup notes, see `license_server/README.md`. License metadata can restrict allowed app company profiles with `allowed_profiles`, `profiles`, `company_profiles`, `allowed_companies`, or `companies`.

## Run From Source

- React/FastAPI web app: `run_react_app.bat`
- Desktop wrapper from source: `run_react_native_app.bat`
- License admin: `run_license_server_admin.bat`

## Key Paths

- Project map: `docs/PROJECT_MAP.md`
- Fast change guide: `docs/FAST_CHANGE_GUIDE.md`
- Active React UI: `react_frontend/src/vietmax/VietmaxApp.tsx`
- Inventory allocation UI: `react_frontend/src/vietmax/InventoryAllocationStage.tsx`
- FastAPI backend: `web_api.py`
- Shared backend logic: `app.py`
- Desktop wrapper: `web_desktop_app.py`
- PyInstaller spec: `ProductCodeFormatterWeb.spec`
- Packaged app output: `deploy/ProductCodeFormatterWeb.exe`

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
