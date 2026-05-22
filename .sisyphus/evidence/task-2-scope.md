# Task 2 Scope Evidence

Confirmed hotspots changed:
- `refreshDerivedCodeViews()` now skips Cao Thành-only price and misorder recomputation for other profiles.
- `productCodeFor()` / `productBaseCode()` now share a per-profile/product base-code cache, reducing repeated `buildCodePreview()` calls across price, misorder, long-code, and modal refresh paths.
- `updateLongCodeCountsChunked()` was added for config operations so long-code recomputation yields between chunks.

Confirmed candidates not changed:
- `refreshNearPhraseGroups()` remains modal-triggered; baseline profile switching did not open the near-phrase modal.
- Backend `/api/config` handlers were not changed because baseline profile-selection lag was local Angular recomputation.
