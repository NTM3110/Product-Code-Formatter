# Refactor Roadmap

The current app works, but `app.py`, `web_api.py`, and
`react_frontend/src/vietmax/VietmaxApp.tsx` are too large. This roadmap keeps
future changes faster without a risky rewrite.

## Phase 1 - Navigation And Guard Rails

Status: started.

- Add `docs/CHANGE_MAP.md` so common user requests map to files/functions.
- Add `docs/PROFILE_EXTENSION_GUIDE.md` for new profiles/workflows.
- Add `product_code/workflows/registry.py` as a small metadata home for
  workflow/profile ownership.
- Keep behavior unchanged.

## Phase 2 - Pure Backend Modules

Goal: move pure helpers out of `app.py` with tests.

Safe candidates:

- product-code sanitization and text normalization
- FAST export row mapping
- inventory pair normalization/application
- review-row section classification

Rules:

- Move one helper family at a time.
- Keep compatibility imports in `app.py` until callers are migrated.
- Run `test_process_workbook.py` after each move.

## Phase 3 - API Routers

Goal: reduce `web_api.py` search time.

Target structure:

```text
api/
  vietmax.py
  inventory.py
  fast.py
  license.py
```

Rules:

- Keep request/response shapes unchanged.
- Move one route group at a time.
- Keep `/api/...` paths identical.

## Phase 4 - React Stage Components

Goal: reduce `VietmaxApp.tsx` churn.

Target structure:

```text
react_frontend/src/vietmax/stages/
  UploadStage.tsx
  ColumnStage.tsx
  CompanyStage.tsx
  ReviewStage.tsx
  MatchStage.tsx
  ProcessStage.tsx
  FastExportStage.tsx
```

Rules:

- Extract presentational parts first.
- Keep state ownership in `VietmaxApp.tsx` until the component boundary is
  stable.
- Run `npm run build` after each extraction.

## Phase 5 - New Workflow Shells

Goal: support profiles that are not product-code workflows.

Future profile example:

- `boc_tach_du_toan`

Target:

- independent backend package
- independent React folder
- shared profile/license shell only

This avoids forcing estimate-extraction screens into the Vietmax stage model.
