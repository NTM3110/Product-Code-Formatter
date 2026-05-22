# Task 3 Scope Evidence

Changed handlers confirmed by Task 1 as custom-configuration lag paths:
- `onProfileChange()` now yields before heavy recomputation and runs derived refresh in chunked mode.
- `importConfig()` and `deleteProfileCache()` now use the same chunked derived refresh path after applying config to loaded companies.
- `saveProfileConfig()` and `exportConfig()` now use operation state so delayed loading can apply if they exceed 2 seconds.

Skipped handlers:
- General upload/check/process handlers were not changed; they are outside the custom-configuration workflow.
