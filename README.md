# Product Code Formatter - MST All Names V6

## Overview

Desktop Flask + Angular tool for formatting `Mã VT` from Excel invoice data.

Main capabilities:
- group sellers by `MST`
- suggest and validate company prefixes
- let users include or skip companies and products
- generate product codes by profile (`Sơn Phương`, `Cao Thành`, `Quang Thịnh`)
- persist config, rules, overrides, price-group settings, and skipped items
- export a processed Excel workbook

## Current Notes

- The company check page shows all company names found under the same `MST`.
- The app still groups by seller `MST`.
- Prefix suggestion uses the most common company name for that `MST`.
- Users can tick or untick which company groups and products to process.
- Rows where quantity is empty, invalid, or `0` are skipped.

## Build

1. Extract the project ZIP.
2. Run `build_exe_auto.bat`.
3. Launch `deploy\ProductCodeFormatter.exe`.

## Run From Source

- Backend: `run_python_app.bat`
- Frontend dev server: in `frontend`, run `npm start`

## Key Paths

- Backend: `app.py`
- Frontend app: `frontend/src/app/app.component.ts`
- Frontend template: `frontend/src/app/app.component.html`
- Frontend styles: `frontend/src/app/app.component.css`
- Frontend tests: `frontend/src/app/app.component.spec.ts`
- Build script: `build_exe_auto.bat`
- Packaged app output: `deploy/ProductCodeFormatter.exe`

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
