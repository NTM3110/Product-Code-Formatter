# Learnings

## 2026-05-20 Task 1 Baseline

- Representative fixture: `uploads/1b6bdcbffa9241339a2494a8b9d97960.xlsx`.
- Loaded data size: 15,634 rows, 1 company, 3,204 products.
- Confirmed profile selection and verify-prefix lag is dominated by Angular main-thread derived-data recomputation, not API latency.
- Confirmed profile selection to Cao Thành took 2,933 ms; verify prefixes took 3,515 ms.
- Implementation should optimize `verifyPrefixes()` / `refreshDerivedCodeViews()` and avoid repeated code preview generation across price, misorder, and long-code refreshes.

## 2026-05-20 Implementation

- Profile switching primary acknowledgement improved to ~117-119 ms on the 3,204-product fixture by yielding before heavy work and chunking long-code recomputation.
- Cao Thành-only price/misorder recomputation is now skipped for other profiles.
- Product base-code previews are cached within the current profile/config state and invalidated on config/rule/prefix changes.
- Delayed config loading shows after 2 seconds for slow config operations and clears on success/failure/superseded operations.
