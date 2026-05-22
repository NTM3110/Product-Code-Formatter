# Task 1 Baseline Evidence

Timestamp: 2026-05-20T03:35:00Z

## Fixture

- File: `uploads/1b6bdcbffa9241339a2494a8b9d97960.xlsx`
- Browser-served app: `http://127.0.0.1:5000/`
- API smoke: `GET /api/config` returned HTTP 200.
- After upload/check: Step 3 loaded with 15,634 rows, 1 company, and 3,204 selected products.

## Baseline Timings

- Upload file to Step 2: 11,466 ms
- Check companies to Step 3: 19,210 ms
- Profile select Sơn Phương -> Cao Thành: 2,933 ms
- Profile select Cao Thành -> Sơn Phương: 1,588 ms
- Export configuration click: 1,632 ms
- Verify prefixes click: 3,515 ms

## Classification

- Confirmed custom-configuration lag source: Angular main-thread derived-data recomputation after profile/config actions.
- Confirmed hotspots eligible for implementation:
  - `verifyPrefixes()` triggers `sortCompanies()` and `refreshDerivedCodeViews()`.
  - `refreshDerivedCodeViews()` recomputes price groups, misorder groups, and long-code counts.
  - `buildPriceConflictRows()`, `refreshMisorderGroups()`, and `updateLongCodeCounts()` repeatedly generate product code previews across 3,204 products.
  - Template calls repeated methods during change detection: `selectedProfileLabel()`, `selectedProfileNote()`, `wordRuleCount()`, `skippedCompanies()`, `productRowSummary()`, and `productPriceSummary()`.
- Backend/API is not the dominant source for profile selection and verify-prefix clicks because those actions are local Angular recomputation after data is already loaded.

## Notes

- Save/delete/import configuration were not executed during baseline because they mutate persisted configuration/cache. Their code paths share the confirmed profile/apply/verify recomputation path and should receive the same delayed-operation handling.
