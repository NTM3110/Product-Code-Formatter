# Product Code Formatter

Product Code Formatter is a Windows-first Excel processing application for Vietnamese invoice and inventory workflows. The desktop application runs a React 19 interface inside pywebview, talks to a local FastAPI server, and delegates workbook rules to Python modules built on pandas and openpyxl.

This README describes the current repository. For implementation rules and a request-to-code lookup table, read [AGENTS.md](AGENTS.md).

## What the application does

The profile selected in the UI determines the available workflow:

| Profile | Current workflow |
|---|---|
| vietmax | Purchase and sales formatting, product review, purchase/sales matching and conversion, inventory allocation, reports, and FAST export |
| son_phuong | Two-phase purchase/sales formatting plus steel-aware inventory allocation and reporting |
| cao_thanh | Generic formatter with price grouping/filtering |
| quang_thinh | Generic product-code formatter |
| viet_hung | Generic product-code formatter |
| ho_guom | Estimate extraction or generic formatter mode |

The Inventory Allocator is also available as a standalone Flask module under inventory_allocation_app/, but its active desktop UI is integrated into the React workflow.

## Runtime architecture

~~~text
react_frontend/src/main.tsx
  -> react_frontend/src/vietmax/VietmaxApp.tsx
  -> react_frontend/src/api.ts
  -> web_api.py (FastAPI)
  -> app.py / product_code/* / workflows/* / inventory_allocation_app/app.py
  -> processed Excel files, reports, and FAST workbooks

web_desktop_app.py
  -> starts FastAPI on localhost
  -> opens the React bundle in pywebview
  -> exposes the native Save As dialog to React
~~~

The active frontend is react_frontend/. There is no active Angular or PySide frontend in this repository.

## Start from source

The preferred acceptance workflow is the isolated runtime preview:

~~~powershell
.\run_runtime_preview.bat
~~~

It runs Vite, FastAPI, and pywebview from source. Preview data is isolated under %LOCALAPPDATA%\ProductCodeFormatterPreview so it does not overwrite the installed application's state.

To recreate preview data from the installed application:

~~~powershell
.\run_runtime_preview.bat -ResetData
~~~

Other launchers:

~~~powershell
.\run_react_app.bat          # FastAPI plus the built React bundle
.\run_react_native_app.bat   # desktop wrapper from source
.\run_license_server_admin.bat
~~~

Frontend development only:

~~~powershell
Set-Location .\react_frontend
npm run dev
~~~

## Verify changes

React or TypeScript changes:

~~~powershell
Set-Location .\react_frontend
npm run build
~~~

From the repository root, run the main Python syntax and workbook regressions:

~~~powershell
.\.venv\Scripts\python.exe -m py_compile app.py web_api.py web_desktop_app.py test_process_workbook.py
.\.venv\Scripts\python.exe -m unittest test_process_workbook.py
~~~

Run focused suites when their owner changes:

~~~powershell
.\.venv\Scripts\python.exe -m unittest test_estimate_extractor.py
.\.venv\Scripts\python.exe -m unittest test_workflow_runtime.py
.\.venv\Scripts\python.exe -m unittest test_release_store.py
.\.venv\Scripts\python.exe -m unittest discover inventory_allocation_app
~~~

See the test matrix in [AGENTS.md](AGENTS.md) before choosing a smaller verification set.

## Build and release

Build the React bundle, Python application, and PyInstaller onedir package:

~~~powershell
.\build_app.ps1 -Version "0.4.0" -Notes "Local build" -Channel dev
~~~

The executable is created at:

~~~text
dist/ProductCodeFormatter/ProductCodeFormatter.exe
~~~

The onedir executable above must remain beside its `_internal` directory. To create the
single-file standalone executable that can be copied by itself to another computer, run:

~~~powershell
.\build_standalone.ps1 -Version "0.4.5" -Notes "Standalone build" -Channel dev
~~~

The standalone build is written to `dist/standalone/ProductCodeFormatter.exe` and copied
to `deploy/ProductCodeFormatter.exe`.

Create a Velopack test or stable release only after the source preview is accepted:

~~~powershell
.\build_release.ps1 -Version "0.4.0" -Notes "Release notes" -Channel test
.\build_release.ps1 -Version "0.4.0" -Notes "Release notes" -Channel stable
~~~

Release feeds are written under Releases/<channel>/; distributable setup files and bundles are copied to deploy/.

## Important paths

| Path | Responsibility |
|---|---|
| react_frontend/src/vietmax/VietmaxApp.tsx | Main application state, stage orchestration, profile workflows, config payloads |
| react_frontend/src/vietmax/workflowStages.ts | Profile capabilities and stage definitions |
| react_frontend/src/vietmax/InventoryAllocationStage.tsx | Allocation configuration, review, reports, and export UI |
| react_frontend/src/estimate/EstimateExtractorWorkflow.tsx | Ho Guom estimate-extraction UI |
| react_frontend/src/api.ts | Typed browser-to-FastAPI calls |
| web_api.py | Active FastAPI endpoints, workflow jobs, uploads, caches, and integration adapters |
| app.py | Shared configuration, product-code, workbook, Vietmax, inventory-pair, and FAST logic |
| product_code/ | Extracted config, license, Excel I/O, updater, release, and workflow-runtime helpers |
| inventory_allocation_app/app.py | Allocation engine, ledger/report generation, and workbook export |
| workflows/estimate_extractor/logic.py | Estimate analysis and export engine |
| web_desktop_app.py | Desktop bootstrap and native file-save bridge |
| ProductCodeFormatterWeb.spec | PyInstaller package definition |
| build_app.ps1 / build_release.ps1 | Application and Velopack release builds |

## Runtime data

Installed application state is stored under %LOCALAPPDATA%\ProductCodeFormatter:

- product_code_config.json - normalized profile settings and saved rules
- license.json - local machine activation state
- default_form_mappings.json and templates/ - persisted form mappings/templates
- sessions/ - resumable workflow session metadata
- uploads/ and outputs/ - runtime workbook artifacts

Do not edit runtime config by hand for feature work. Use the config helpers in app.py and the API save endpoints so normalization and backups continue to work.

## Security and generated files

- License activation is machine-bound and validated through the Keygen client in product_code/license_client.py.
- Never place admin tokens, private certificates, license keys, or .env secrets in application code or commits.
- Do not hand-edit build/, dist/, deploy/, uploads/, outputs/, or __pycache__/ during normal implementation.
- Output workbook rules and user configuration are business data; preserve backward compatibility when changing their shapes.
