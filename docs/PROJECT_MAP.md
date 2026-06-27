# Project Map

This repository has a few large files, so start here before editing.

## Active Product

Product Code Formatter desktop app:

- `react_frontend/src/vietmax/VietmaxApp.tsx` - active React wizard UI.
- `react_frontend/src/vietmax/InventoryAllocationStage.tsx` - inventory allocation UI, report view, export stage.
- `react_frontend/src/api.ts` - frontend HTTP client.
- `react_frontend/src/types.ts` - shared frontend data shapes.
- `react_frontend/src/styles.css` - active UI styling.
- `web_api.py` - FastAPI API used by React and desktop app.
- `app.py` - shared Excel/business logic used by `web_api.py`.
- `web_desktop_app.py` - pywebview desktop wrapper.
- `ProductCodeFormatterWeb.spec` - PyInstaller build spec.

## Secondary Product

Inventory Allocator:

- `inventory_allocation_app/app.py` - allocation engine and report generation.
- `inventory_allocation_app/test_inventory.py` - allocator tests.
- `inventory_allocation_app/company_detection_rules.json` - company/profile detection rules.

Only edit this area for Stage 12/report/allocation work.

## Legacy Areas

- `frontend/` - old Angular UI. Do not edit unless the task explicitly targets legacy UI.
- `static/` - old built Angular output.

## Runtime And Generated Areas

Avoid editing these by hand:

- `uploads/`
- `outputs/`
- `build/`
- `dist/`
- `deploy/` except when copying a finished `.exe`
- `__pycache__/`
- `.pycompile_check*/`
- `.playwright-mcp/`
- `.sisyphus/`

## Common Edit Targets

Fast task routing:

- `docs/AI_START_HERE.md` - first file to read after `AGENTS.md` in future AI sessions.
- `docs/CHANGE_MAP.md` - user request -> exact file/function owner.
- `docs/PROFILE_EXTENSION_GUIDE.md` - how to add a company profile or a new workflow.
- `docs/REFACTOR_ROADMAP.md` - safe order for future code extraction.
- `product_code/workflows/registry.py` - lightweight workflow/profile ownership map.

Prefix/company selection:

- React state and UI: `react_frontend/src/vietmax/VietmaxApp.tsx`
- Backend load/save: `web_api.py`
- Config normalization: `app.py`
- Data type shape: `react_frontend/src/types.ts`

Review Ma VT:

- UI tables/buttons: `react_frontend/src/vietmax/VietmaxApp.tsx`
- API endpoint: `web_api.py` `/api/vietmax/review`
- Core comparison logic: `app.py` `vietmax_product_review_rows`

Khớp mua vào/bán ra:

- UI: `react_frontend/src/vietmax/VietmaxApp.tsx`
- API endpoint: `web_api.py` `/api/vietmax/matches`
- Core logic: `app.py` `build_vietmax_ban_ra_purchase_matches`

Create processed purchase/sales files:

- UI: `react_frontend/src/vietmax/VietmaxApp.tsx`
- API endpoint: `web_api.py` `/api/vietmax/process`
- Core logic: `app.py` `process_workbook`

Stage 12 inventory allocation:

- UI: `react_frontend/src/vietmax/InventoryAllocationStage.tsx`
- API endpoints: `web_api.py` `/api/inventory-allocation/*`
- Core engine: `inventory_allocation_app/app.py`

License/admin:

- Desktop license client logic: `app.py`
- License API wrapper: `web_api.py`
- Admin app: `license_server_admin.py`

## Config Flow

Main config file:

- Development workspace: `product_code_config.json`
- Packaged desktop runtime: `%LOCALAPPDATA%\ProductCodeFormatter\product_code_config.json`

Do not edit config directly for normal changes. Use:

- `app.py` `load_config()`
- `app.py` `save_config()`
- `app.py` `normalize_config()`
- `app.py` `empty_profile_config()`

If a new persisted field is added, add it to:

1. `empty_profile_config()` in `app.py`
2. `normalize_profile_config()` in `app.py`
3. API save/load in `web_api.py`
4. TypeScript type/state in `react_frontend/src/types.ts` or `VietmaxApp.tsx`
