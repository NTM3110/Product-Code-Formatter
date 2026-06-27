# Profile And Workflow Extension Guide

This guide is for adding new company profiles without making `app.py` and
`VietmaxApp.tsx` harder to change.

## Choose The Workflow Type First

Use an existing workflow only when the new company follows the same business
shape.

- Product code formatter workflow:
  - Reads invoice workbooks.
  - Generates or reviews Ma VT.
  - Exports processed FDI/FAST files.
- Inventory allocation workflow:
  - Starts from processed purchase/sales files.
  - Builds stock/sales reports.
- Estimate extractor workflow (`boc tach du toan`, future):
  - Reads construction/estimate workbooks.
  - Extracts work items/material/labor/machine/cost columns.
  - Should be separate from Product Code Formatter stages.

## Add A Product Code Formatter Company

1. Add profile identity:
   - `app.py`
     - `PROFILE_LABELS`
     - `PROFILE_ALIASES`
     - `empty_profile_config`
   - `license_server_admin.py`
     - product/profile options if needed
   - React profile dropdown source in `react_frontend/src/vietmax/VietmaxApp.tsx`

2. Add backend rules:
   - Prefer small helper functions near the related profile section.
   - Keep config fields profile scoped.
   - If the profile has separate purchase/sales behavior, store config in
     separate purchase and sales scopes.

3. Add API endpoints only when needed:
   - Reuse Vietmax endpoints only if the request/response shape is the same.
   - Add a profile branch in `web_api.py` only at the boundary layer.
   - Keep workbook/business logic in `app.py` or a future extracted module.

4. Add React UI:
   - Reuse existing stages if the workflow is the same.
   - If the profile has a truly different sequence, make a profile-specific
     workflow component instead of adding many conditionals to every Vietmax
     stage.

5. Add tests:
   - Code generation tests in `test_process_workbook.py`.
   - Config persistence tests when adding fields.
   - Workbook processing tests with the smallest possible sample fixture.

## Add A New Non-Product Workflow

Do not put a new workflow like `boc tach du toan` inside the Vietmax stage
machine. Use this structure instead:

1. Add workflow metadata to `product_code/workflows/registry.py`.
2. Add a backend module or package for the workflow logic, for example:
   - `estimate_extraction/`
   - `estimate_extraction/app.py`
   - `estimate_extraction/tests/`
3. Add FastAPI routes under a separate prefix, for example:
   - `/api/estimate-extraction/*`
4. Add a React component folder:
   - `react_frontend/src/estimate/`
5. Let the desktop shell choose the workflow from the selected profile.
6. Keep license metadata profile-based so admin can enable/disable the new
   profile without changing client code.

## Config Rules

Every persisted field must be added in all relevant places:

1. `app.py` `empty_profile_config`
2. `app.py` `normalize_profile_config`
3. `web_api.py` request/response model or save/load endpoint
4. React state/type
5. Tests

Do not store runtime-only values in config:

- uploaded file paths
- output save paths
- progress/job ids
- generated temp/cache filenames

Do store user decisions:

- selected columns
- selected companies/products
- prefix strategy/manual values
- review pair decisions
- split/merge decisions
- word replacement rules
- inventory pair rules
- FAST/report mapping options

## Recommended Extraction Order

When refactoring, use this order to reduce risk:

1. Extract pure helpers with tests.
2. Extract profile/workflow metadata.
3. Extract FAST export helpers.
4. Extract Vietmax review/match helpers.
5. Extract React stage components.
6. Only then split large endpoint files.

Avoid moving UI and backend logic in the same patch unless the behavior change
requires it.
