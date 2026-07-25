# Product Code Formatter - Agent Guide

This is the durable implementation guide for future coding sessions. It was verified against the working tree on 2026-07-16. Use symbol names and API paths below as search anchors; line numbers are intentionally omitted because the large workflow files change frequently.

## Start here

Read in this order:

1. README.md for product, runtime, build, and release usage.
2. This file for ownership and change routing.
3. The exact owner file and symbol from the request-to-code map below.
4. The closest focused test before editing behavior.

The code is the final source of truth. If a guide conflicts with the current implementation, verify current behavior, update the guide in the same change, and preserve backward compatibility unless the request explicitly changes it.

## Product boundary

The active application is:

- React 19 + Vite in react_frontend/.
- FastAPI in web_api.py.
- Shared formatter and workbook logic in app.py and product_code/.
- A pywebview desktop host in web_desktop_app.py.
- A PyInstaller onedir package followed by Velopack packaging.

Additional engines integrated into the same UI:

- Inventory allocation: inventory_allocation_app/app.py.
- Ho Guom estimate extraction: workflows/estimate_extractor/logic.py.

Do not recreate or modify removed Angular, PySide, or old static frontend stacks. The repository has one active frontend: React.

## End-to-end ownership

~~~text
UI event and state
  react_frontend/src/vietmax/VietmaxApp.tsx
  react_frontend/src/vietmax/*.tsx
  react_frontend/src/estimate/EstimateExtractorWorkflow.tsx
        |
HTTP contract
  react_frontend/src/api.ts + react_frontend/src/types.ts
        |
API orchestration, upload/cache/job handling
  web_api.py + product_code/workflow_runtime.py
        |
Business/workbook engines
  app.py
  inventory_allocation_app/app.py
  workflows/estimate_extractor/logic.py
        |
Excel/report artifacts
~~~

Keep business rules out of React. Keep request parsing, HTTP responses, and job orchestration in web_api.py. Put reusable formatter/workbook rules in app.py or an extracted Python module, and keep allocation or estimate rules in their existing engines.

## Current profiles and workflow ownership

The UI profile list and stage capabilities live in react_frontend/src/vietmax/workflowStages.ts.

| Profile key | Workflow owner |
|---|---|
| vietmax | Native two-phase workflow in VietmaxApp; matching, allocation, and FAST enabled |
| son_phuong | Native two-phase workflow plus steel-aware allocation |
| cao_thanh | Generic formatter with CaoThanhPriceStage |
| quang_thinh | Generic formatter |
| viet_hung | Generic formatter |
| ho_guom | EstimateExtractorWorkflow in estimate mode; generic formatter shell in formatter mode |

Backend profile aliases and effective Vietmax purchase/sales keys are owned by PROFILE_LABELS, PROFILE_ALIASES, profile_key, and effective_processing_profile in app.py. The lightweight registry in product_code/workflows/registry.py is descriptive; it does not replace the React stage definitions.

## Request-to-code map

### Application shell, profiles, or stage navigation

- Profile keys, capabilities, and stages: react_frontend/src/vietmax/workflowStages.ts - profiles, profileCapabilities, stagesForProfile, sonPhuongProfileStages.
- Stage groups/buttons: react_frontend/src/vietmax/StageNavigation.tsx - StageNavigation, StageGroup.
- Stage guards, transitions, reset/resume, and render ownership: react_frontend/src/vietmax/VietmaxApp.tsx - VietmaxApp, canEnterStage, goToStage, renderTwoPhaseStage, renderProfileStage.
- Shared stage panels: react_frontend/src/vietmax/basicStages.tsx.
- Global styling: react_frontend/src/styles.css.

When adding a profile, update the TypeScript ProfileKey, profile list, capabilities/stages, initial workflow state, config handling, backend profile labels/aliases, license metadata handling, and tests together.

### Upload, workflow resume, background jobs, or cached outputs

- Frontend calls: react_frontend/src/api.ts - uploadExcel, getWorkflowSession, startWorkflowProcessJob, waitForWorkflowJob.
- API routes: web_api.py - /api/files/upload, /api/session/current, /api/session/close, /api/workflow/process-jobs, /api/workflow/json-jobs, /api/jobs/{job_id}, /api/artifacts/{artifact_id}.
- Session/job implementation: product_code/workflow_runtime.py - WorkflowSessionStore, WorkflowJobManager, stable_signature, file_sha256.
- Uploaded workbook conversion: product_code/excel_io.py - uploaded_workbook_content_for_openpyxl, workbook_content_for_openpyxl.

Cache signatures must include every input that changes the output. Cached files are runtime artifacts, not configuration.

### Form mapping and processing groups

- UI: VietmaxApp.tsx - FormatMappingStage, FormatScopeModal, FormatGroupModal, TransformRuleEditor.
- Default mapping API: web_api.py - /api/vietmax/format-mapping-defaults.
- Persisted mapping normalization: app.py - empty_profile_config, normalize_profile_config, load_system_default_form_mappings.
- FAST application of mappings: app.py - fast_merge_mapping_forms, fast_import_sheet_rows, fast_import_multi_sheet_workbook.

Purchase and sales scopes are separate. A mapping may explicitly target purchase, sales, or both; do not infer a different scope in the exporter.

### Company list, prefixes, included products, or missing MST

- UI: VietmaxApp.tsx - CompanyRulesStage, CompanyGroupRows, MissingMstCompanyRows and the apply company/product handlers.
- Analysis API: web_api.py - /api/check, /api/vietmax/analyze.
- Core grouping: app.py - company_rows, missing_mst_company_warnings, choose_company, analyze.
- Prefix rules: app.py - remove_company_location_phrases, suggest_prefix, compute_prefix_strategies, prefix_last_n_words, prefix_2_words_mst, prefix_all_name_words.

Pending checkboxes and edits must not change the applied company/product selection until the explicit apply action runs. Purchase and sales prefix/config scopes remain independent.

### Product-code generation, cleanup, or manual replacement

- Core code path: app.py - make_product_part, make_code, sanitize_product_code, code_words, apply_word_rules_to_words, remove_repeated_phrases, apply_product_code_replacement.
- Extracted sanitizer helper: product_code/code_normalization.py.
- Preview APIs: web_api.py - /api/product-preview, /api/vietmax/product-preview.
- Frontend payloads: VietmaxApp.tsx - loadGenericProductPreviewCodes, buildGenericProcessPayload, buildGenericSalesProcessPayload, buildPurchaseProcessPayload, buildSalesProcessPayload.
- Regression owner: test_process_workbook.py.

MAX_CODE_LENGTH in app.py is the shared warning threshold. Preserve profile-specific rules and sanitize final codes at the business-logic boundary, not only in the UI.

### Product review and same/near-similar names

- UI: VietmaxApp.tsx - ReviewStage, ReviewTable, ReviewGroupRows, buildReviewRules, compactReviewRule.
- APIs: web_api.py - /api/review, /api/vietmax/review.
- Core comparison/persistence: app.py - vietmax_product_review_rows, normalize_vietmax_internal_merges, apply_vietmax_internal_merges_to_products, vietmax_internal_merge_lookup_key.

Review settings are phase- and comparison-scope-sensitive. Do not write review choices into unrelated word-rule or manual-override fields.

### Vietmax purchase/sales matching or unit conversion

- UI: VietmaxApp.tsx - MatchStage, MatchTable, ConversionFormulaInputs, buildSalesMatchRules.
- Frontend API: react_frontend/src/api.ts - createSalesMatches, exportMatches.
- API routes: web_api.py - /api/vietmax/sales-match, /api/vietmax/export-matches.
- Core matching: app.py - build_vietmax_ban_ra_purchase_matches, build_vietmax_khh_exact_purchase_matches, vietmax_purchase_products_from_workbook, vietmax_ban_ra_sales_products_from_workbook.
- Conversion: app.py - parse_vietmax_conversion_formula, vietmax_conversion_quantity_factor, apply_vietmax_match_conversion_to_row, preserve_vietmax_conversion_settings.

Only confirmed rows receive a purchase-code override or conversion. Conversion modes remain explicit and safely parsed; never evaluate arbitrary expressions.

### Generic/Vietmax workbook processing and output files

- Frontend actions and payloads: VietmaxApp.tsx process/download handlers and process-payload builders.
- Frontend APIs: api.ts - processGenericWorkbookCache, processVietmaxPurchase, downloadCachedFile.
- API routes: web_api.py - /api/process, /api/vietmax/process, /api/files/download/{saved_name}.
- Core engine: app.py - process_workbook, resolve_output_path, read_workbook, make_excel_workbook.

Processing should create/reuse a cache once for a stable input signature and use the desktop Save As bridge for the user's final destination. If an output is locked by Excel, surface a close-and-retry message.

### Inventory pairs (Ma kho and TK vat tu)

- UI: VietmaxApp.tsx - InventoryPairEditor and scoped inventory update helpers.
- Config normalization: app.py - normalize_inventory_pairs, normalize_inventory_pair_rules, empty_profile_config, normalize_profile_config.
- Workbook application: app.py - apply_inventory_pairs, inventory_column_indexes, processed_inventory_column_indexes.

Treat the warehouse code/account as one pair. Purchase, sales, and generic scopes must not overwrite each other.

### Inventory allocation, Son Phuong steel rules, or reports

- Active UI: react_frontend/src/vietmax/InventoryAllocationStage.tsx - InventoryAllocationStage, SonPhuongSalesPairEditor, SonPhuongAllocationReviewStage, InventoryAllocationOverviewStage, InventoryAllocationReportStage, InventoryAllocationExportStage.
- Frontend API: api.ts - startInventoryAllocation, getInventoryAllocationJob, createSonPhuongProcessedSales, downloadInventoryAllocationReport.
- Integration routes: web_api.py - /api/inventory-allocation/analyze-job, /api/inventory-allocation/analyze-job/{job_id}, /api/inventory-allocation/analyze-job/{job_id}/create-sales-fdi, /api/inventory-allocation/download/{job_id}.
- Allocation engine: inventory_allocation_app/app.py - read_lines, clean_policy, allocate_stock, build_inventory_ledger, report_view_for_ui, create_output_workbook.
- Son Phuong matching: the son_phuong_*, steel_*, barem, reservation, and allocation helpers in inventory_allocation_app/app.py.
- Company detection persistence: inventory_allocation_app/company_detection_rules.json through build_company_detection_rules and save_company_detection_rules.

Processed purchase and sales workbooks are the allocation source of truth. Do not reapply formatter conversion rules in the allocator. Keep optional opening inventory optional and UI/export date filtering consistent.

### Cao Thanh price filtering

- UI/state: VietmaxApp.tsx - CaoThanhWorkflow, CaoThanhPriceStage, caoThanhRangeRules.
- API: web_api.py - /api/export_price_report, /api/process.
- Core/config: app.py - normalize_price_group_rules, price_ranges_from_groups, raw_price_group, price_group_suffix, merge_price_ranges.
- Tests: Cao Thanh cases in test_process_workbook.py.

Price grouping is based on the final customized product code. Preserve skipped products as skipped and keep the price stage between review and export.

### Ho Guom estimate extraction

- UI: react_frontend/src/estimate/EstimateExtractorWorkflow.tsx.
- Frontend API: api.ts - uploadEstimateWorkbook, analyzeEstimateWorkbook, exportEstimateWorkbook.
- API routes: web_api.py - /api/estimate/upload, /api/estimate/analyze, /api/estimate/export.
- Engine: workflows/estimate_extractor/logic.py - list_estimate_workbook_sheets, analyze_estimate_workbook, create_estimate_output_workbook.
- Tests: test_estimate_extractor.py.

Keep estimate extraction separate from product-code workbook logic even though both modes are selectable under the Ho Guom profile.

### FAST import/export

- UI: react_frontend/src/vietmax/basicStages.tsx - FastImportExportStage.
- Frontend API: api.ts - validateFastImportProcessedFile, createVietmaxFastImportPackage.
- API routes: web_api.py - /api/vietmax/validate-fast-import-processed, /api/vietmax/fast-import-package.
- Core: app.py - validate_fast_import_processed_dataframe, fast_import_sheet_rows, split_fast_purchase_invoice_rows, fast_import_multi_sheet_workbook.

Form mappings decide which source phases are required. Purchase and sales processed files remain the inputs; preserve invoice splitting and sheet naming rules.

### Configuration persistence or import/export

- Defaults and normalization: app.py - empty_profile_config, normalize_profile_config, normalize_config.
- Durable storage: app.py - load_config, save_config; file helpers in product_code/config_store.py.
- API: web_api.py - /api/config, /api/config/profile/{profile}, /api/vietmax/save-config, /api/vietmax/import-config/{phase}.
- React payloads: VietmaxApp.tsx - buildConfigPayloads, buildConfigPayload, buildSalesConfigPayload, buildVietmaxConfigExportSnapshot.

For every new persisted field, update defaults, normalization/migration, API handling, TypeScript state/types, import/export snapshots when applicable, and round-trip tests. Do not edit product_code_config.json directly as an implementation shortcut.

### License, allowed profiles, or activation

- UI/state: VietmaxApp.tsx - licenseAllowsDropdownProfile, licenseAllowsSelectedProfile, activation toolbar.
- Frontend API: api.ts - getLicenseStatus, activateLicense, reloadLicense.
- API routes: web_api.py - /api/license/status, /api/license/activate, /api/license/reload.
- Shared facade/config: app.py - license_allows_profile, license_has_local_activation, configured_license_server_urls.
- Keygen client: product_code/license_client.py - local_machine_fingerprint, keygen_validate_license, keygen_activate_machine, activate_keygen_license, public_license_status.
- Server-side admin tool: license_server_admin.py.

Never embed Keygen admin tokens or secrets in the desktop client. Public HTTP is restricted to local/private hosts by the Keygen URL validation helpers. Preserve machine-bound activation and legacy fingerprint compatibility.

### Application update, build, or release

- UI and calls: VietmaxApp.tsx update modal; api.ts - checkForUpdate, applyUpdate.
- API routes: web_api.py - /api/update/check, /api/update/apply.
- Update client: product_code/update_client.py.
- Embedded build metadata: product_code/release_version.py and ProductCodeFormatterWeb.spec.
- Desktop startup: web_desktop_app.py - run_velopack_startup_hooks, cleanup_legacy_install_artifacts, main.
- Build scripts: build_app.ps1, build_standalone.ps1, build_release.ps1.
- Release tests: test_release_store.py.

ProductCodeFormatterWeb.spec creates the onedir package at dist/ProductCodeFormatter for Velopack. ProductCodeFormatterStandalone.spec creates the self-contained `dist/standalone/ProductCodeFormatter.exe`; use `build_standalone.ps1` whenever the user requests one EXE that can be copied by itself.

### Native desktop window or Save As behavior

- web_desktop_app.py - DesktopApi.save_file, main, and the find_free_port integration.
- React save logic: VietmaxApp.tsx - saveBlob, blobToBase64, downloadBlob.

Keep browser fallback behavior working when the native window.pywebview.api bridge is unavailable.

## Configuration and runtime paths

app.py owns these paths:

- Installed state: %LOCALAPPDATA%\ProductCodeFormatter.
- Config: %LOCALAPPDATA%\ProductCodeFormatter\product_code_config.json.
- License: %LOCALAPPDATA%\ProductCodeFormatter\license.json.
- Templates/default mappings: the same application-data directory.
- Session state: APP_DATA_DIR / sessions in web_api.py.
- Source preview state: %LOCALAPPDATA%\ProductCodeFormatterPreview, established by run_runtime_preview.ps1.

When frozen, BASE_DIR is the application-data directory. During normal source runs it is the repository unless PRODUCT_CODE_FORMATTER_RUNTIME_DIR overrides it. Never hardcode a user's directory.

## Output invariants

- Product-code warning threshold: MAX_CODE_LENGTH = 50.
- Normal processed suffix: _fdi.xls.
- UP companion suffixes: _UP_mua_vao.xls and _UP_ban_ra.xls, built from mau_mua_vao_up.xlsx and mau_ban_ra_up.xlsx (with the legacy sales template as fallback).
- Inventory fields are coupled: Ma kho plus TK vat tu.
- Companion template mappings remain TK vat tu -> AB, Ma kho -> AF, Ma vat tu -> AG unless an explicit request changes them.
- Empty, invalid, or zero quantity rows are not normal processable rows; see should_process_qty.
- Unchecked companies/products/review matches must not receive downstream overrides or conversions.
- Purchase and sales config, review, prefix, inventory-pair, and form-mapping scopes remain separate.

## Verification matrix

Run the smallest adequate set, but always build React after React changes and compile Python after Python changes.

| Changed area | Required verification |
|---|---|
| React/TypeScript | Set-Location react_frontend; npm run build |
| Main or extracted Python | .\.venv\Scripts\python.exe -m py_compile app.py web_api.py web_desktop_app.py test_process_workbook.py |
| Formatter, config, review, match, FAST, or workbook logic | .\.venv\Scripts\python.exe -m unittest test_process_workbook.py |
| Estimate extraction | .\.venv\Scripts\python.exe -m unittest test_estimate_extractor.py |
| Workflow sessions/jobs | .\.venv\Scripts\python.exe -m unittest test_workflow_runtime.py |
| Inventory allocation/report engine | .\.venv\Scripts\python.exe -m unittest discover inventory_allocation_app |
| Release/update scripts or metadata | .\.venv\Scripts\python.exe -m unittest test_release_store.py plus the relevant build script when requested |

Do not declare a React change complete without npm run build, or a Python change complete without syntax compilation. A failing test must be investigated; do not assume an open Excel file unless the error proves a file lock.

## Source preview and packaging policy

After behavior changes, let the user test the isolated source preview before packaging unless they explicitly ask for a build immediately:

~~~powershell
.\run_runtime_preview.bat
~~~

Use -ResetData only when the user wants preview state recreated from the installed application. It deletes only the guarded preview directory and then seeds it again.

Only after acceptance, or on an explicit request, build:

~~~powershell
.\build_app.ps1 -Version "<version>" -Notes "<notes>" -Channel dev
~~~

For a single-file standalone EXE that does not require installation or an `_internal` folder:

~~~powershell
.\build_standalone.ps1 -Version "<version>" -Notes "<notes>" -Channel dev
~~~

Create a Velopack release only when explicitly requested:

~~~powershell
.\build_release.ps1 -Version "<version>" -Notes "<notes>" -Channel test
.\build_release.ps1 -Version "<version>" -Notes "<notes>" -Channel stable
~~~

## Working conventions

- Use PowerShell syntax. Never use && or ||; use separate calls, semicolons, or if ($?).
- Search with rg and stable symbols/API paths.
- Keep fixes scoped; do not combine a production fix with an unrelated large-file refactor.
- Preserve unrelated working-tree changes. Do not use destructive Git commands.
- Do not run git commit, git push, git reset, git rebase, or branch mutations unless explicitly asked.
- Do not hand-edit generated/runtime directories: build/, dist/, deploy/, Releases/, uploads/, outputs/, __pycache__/, .codex_tmp/.
- Do not package a new EXE merely to validate source logic.
- Keep secrets and live license data out of code, docs, logs, and commits.

## Change checklist

Before editing:

1. Identify the profile and phase: generic, purchase, sales, price, allocation, FAST, or estimate.
2. Identify the owning layer: React state/UI, HTTP contract, API orchestration, or Python business logic.
3. Read the current symbol and the closest regression test.
4. Decide whether the value is persisted config, workflow session state, cache signature input, or generated output.

Before handoff:

1. Verify all modified paths and symbols still exist.
2. Run the verification matrix for every changed owner.
3. Review git diff without overwriting unrelated changes.
4. Update README.md or this guide when architecture, commands, profiles, routes, or ownership changed.
5. Report source-preview readiness separately from packaged-release readiness.

## Compact handoff format

When a future session needs a short continuation summary, use:

~~~text
Goal:
- one sentence

Changed Files:
- path - reason

Current Behavior:
- only verified behavior

Open Work:
- exact next steps

Verification:
- latest commands and results

Known Constraints:
- only constraints relevant to the open work
~~~
