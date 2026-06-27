# Change Map

Use this file to jump to the right owner for common requests. It is intentionally
practical: user wording first, files/functions second.

## Product Code And Review

User asks about:

- Ma VT generation
- removing special characters
- case/accent/space normalization
- "cung form ten hang"
- saved review pairs
- split same Ma VT in sales review

Edit/check:

- `app.py`
  - `make_product_part`
  - `make_code`
  - `sanitize_product_code`
  - `vietmax_product_review_rows`
  - review persistence helpers near Vietmax config code
- `web_api.py`
  - Vietmax review endpoints
  - config save/load for review sections and scopes
- `react_frontend/src/vietmax/VietmaxApp.tsx`
  - `ReviewStage`
  - review grouping/filter controls
  - apply/save review handlers
- Tests:
  - `test_process_workbook.py`

Rules to preserve:

- Case/accent/extra-space-only differences are normalized automatically and
  should not create review noise.
- Purchase and sales review config must stay separated.
- Section-scoped review config must load/save by the same section scope.
- Merge selections may persist the chosen side/code, but they must not write
  into generic word replacement/manual-code override fields.

## Company Prefix And Company List

User asks about:

- prefix strategies: 2 words, MST, 2 words + MST
- duplicated prefixes
- company list ordering
- manual prefix edits
- select all / unselect all company rows

Edit/check:

- `react_frontend/src/vietmax/VietmaxApp.tsx`
  - `CompanyRulesStage`
  - prefix strategy state
  - company row ordering and duplicate groups
- `web_api.py`
  - analyze/company apply endpoints
  - config persistence endpoints
- `app.py`
  - company grouping/prefix helper functions
  - `empty_profile_config`
  - `normalize_profile_config`

Rules to preserve:

- Editing prefix fields must not recalculate/reorder the company list until the
  explicit apply button is clicked.
- Purchase and sales prefix config are separate.
- Each prefix strategy has its own manual values.

## Cross-File Purchase/Sales Match

User asks about:

- `Khớp mua vào`
- saved matched rows
- unit conversion formula
- same product with different unit
- processed sales should align unit/quantity/price to purchase when selected

Edit/check:

- `app.py`
  - `build_vietmax_ban_ra_purchase_matches`
  - conversion parsing/apply helpers
  - Vietmax sales processing path inside `process_workbook`
- `web_api.py`
  - `/api/vietmax/matches`
  - match config save/load endpoints
- `react_frontend/src/vietmax/VietmaxApp.tsx`
  - stage 8 match table
  - conversion inputs and save handler

Rules to preserve:

- Only checked/confirmed rows receive purchase-code override or conversion.
- Saved match config is company/profile scoped.
- Processed sales files are the source of truth for later inventory/report
  stages; inventory allocation should not re-apply conversion formulas.

## Processed FDI Files

User asks about:

- stage 5 create purchase file
- stage 11 create sales file
- stuck loading while creating file
- output save dialog
- FDI workbook format

Edit/check:

- `app.py`
  - `process_workbook`
  - `process_zip_stream`
  - output path/file-lock helpers
- `web_api.py`
  - `/api/vietmax/process`
  - progress/job handling
- `react_frontend/src/vietmax/VietmaxApp.tsx`
  - stage 5 and stage 11 render/handlers

Rules to preserve:

- Creating cache should happen once per stage data set.
- Clicking export should reuse the cached output.
- Output location is chosen by the user through the save dialog.
- Do not hardcode user-specific folders.

## Inventory Allocation And Reports

User asks about:

- stage 12 upload processed purchase/sales
- stage 13 report view
- stage 14 export report
- so chi tiet
- bao cao ban hang
- tong hop NXT
- sticky total rows
- date range filtering
- report scroll/performance

Edit/check:

- `inventory_allocation_app/app.py`
  - allocation engine
  - report data generation
  - Excel export
- `web_api.py`
  - `/api/inventory-allocation/*`
  - job/progress endpoints
- `react_frontend/src/vietmax/InventoryAllocationStage.tsx`
  - upload/config stage
  - report view tabs
  - table virtualization/chunked rendering
  - export stage

Rules to preserve:

- Processed purchase and processed sales files are source of truth.
- Purchase inventory-pair rules and sales inventory-pair rules are separate.
- Optional opening inventory must stay optional.
- Date filters must affect UI reports and exported report data consistently.

## FAST Export

User asks about:

- stage 15 `Xuat FAST`
- Hoadonmuahang / Hoadonbanhang / DMvat tu / DMkhachhang
- one workbook with four sheets
- `.xls` export
- duplicate purchase invoice number split

Edit/check:

- `app.py`
  - FAST header constants
  - `fast_import_multi_sheet_workbook`
  - `fast_import_sheet_rows`
  - `split_fast_purchase_invoice_rows`
  - `validate_fast_import_processed_dataframe`
- `web_api.py`
  - FAST import/export endpoint
- `react_frontend/src/vietmax/VietmaxApp.tsx`
  - stage 15 UI

Rules to preserve:

- Only fill sample columns that have values in the sample file.
- Purchase FDI maps to Hoadonmuahang.
- Sales FDI maps to Hoadonbanhang.
- Duplicate purchase `so_ct` across different suppliers must be split into
  `Hoadonmuahang-1`, `Hoadonmuahang-2`, etc., and reported.

## License And Admin

User asks about:

- license activation
- 403 from Keygen
- profile list in license admin
- server URL/IP behavior

Edit/check:

- `app.py`
  - `activate_keygen_license`
  - Keygen validation/machine helpers
  - local license storage
- `web_api.py`
  - public license status/activation endpoints
- `license_server_admin.py`
  - admin UI and metadata update

Rules to preserve:

- Client app needs server URL and license key.
- Admin token/secrets must never be embedded in desktop app.
- Local activation should remain durable on the same machine fingerprint.

## New Company Or New Workflow

For a new Product Code Formatter company profile, start with
`docs/PROFILE_EXTENSION_GUIDE.md`.

For a future non-product workflow such as `boc tach du toan`, do not force it
into Vietmax stages. Add a new workflow shell/profile and separate backend
router/logic, then expose it through the shared desktop shell and license
profile list.
