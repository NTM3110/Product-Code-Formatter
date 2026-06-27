# AI Start Here

Read these in order at the start of a new coding session:

1. `AGENTS.md`
2. `docs/PROJECT_MAP.md`
3. `docs/CHANGE_MAP.md`
4. `docs/FAST_CHANGE_GUIDE.md`

Then search by symbol, not by broad text.

## High-Value Symbols

Backend:

- Config defaults: `empty_profile_config`
- Config normalization: `normalize_profile_config`
- Product-code cleanup: `sanitize_product_code`
- Product-code generation: `make_product_part`, `make_code`
- Review rows: `vietmax_product_review_rows`
- Review merge config: `normalize_vietmax_internal_merges`,
  `apply_vietmax_internal_merges_to_products`
- Purchase/sales matching: `build_vietmax_ban_ra_purchase_matches`
- Inventory pair logic: `normalize_inventory_pairs`, `apply_inventory_pairs`
- Main FDI processing: `process_workbook`
- FAST export: `fast_import_multi_sheet_workbook`,
  `fast_import_sheet_rows`, `split_fast_purchase_invoice_rows`

API:

- React backend: `web_api.py`
- Vietmax endpoints: search `/api/vietmax`
- Inventory endpoints: search `/api/inventory-allocation`
- License endpoints: search `license`
- FAST endpoints: search `fast`

React:

- Main wizard: `react_frontend/src/vietmax/VietmaxApp.tsx`
- Stage rendering: search `renderStage`
- Company/prefix UI: `CompanyRulesStage`
- Review UI: `ReviewStage`
- FAST UI: `FastImportExportStage`
- Inventory UI: `react_frontend/src/vietmax/InventoryAllocationStage.tsx`

## Fast Triage Questions

Before editing, answer these:

- Which profile is affected?
- Is it purchase, sales, inventory, FAST, license, or a future workflow?
- Is the config field user decision data or runtime/cache data?
- Is the bug in UI state, API shape, workbook logic, or exported format?
- Does the fix need React build, Python compile, workbook tests, or EXE rebuild?

## Refactor Rule

Do not split a large file while fixing a production bug unless the split is
needed for that bug. Prefer one of these small moves:

- add a focused helper with tests
- move one pure helper family
- add a compatibility wrapper
- add or update the change map
