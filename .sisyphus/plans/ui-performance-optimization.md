# UI Performance Optimization for Custom Configuration

## TL;DR
> **Summary**: Optimize lag in the Angular custom-configuration workflow by profiling first, reducing confirmed main-thread/render recomputation hotspots, and adding delayed loading feedback for operations exceeding 2 seconds.
> **Deliverables**:
> - Baseline and post-fix timing evidence for custom-configuration actions.
> - Narrow Angular performance fixes in `frontend/src/app/app.component.ts`, `.html`, and `.css` only if profiling confirms need.
> - Delayed loading indicator that appears after 2 seconds, never flashes for fast operations, and clears on success/failure.
> - Focused Angular specs plus build/API/browser QA evidence.
> - Atomic commit(s) preserving existing unstaged and committed behavior.
> **Effort**: Medium
> **Parallel**: YES - 4 waves
> **Critical Path**: Task 1 → Tasks 2/3/4 → Task 5 → Task 6 → Final Verification Wave

## Context
### Original Request
- “Optimise the UI of this project: `C:\Users\shou\Downloads\product_code_flask_app_mst_all_names_v6 (1)` after recent unstaged change and commit. It gets laggy, especially when clicking on the custom configuration.”
- Follow-up: “and add the UI to show loading when the load is more than 2s”

### Interview Summary
- Success target: custom configuration click/update should respond within about 300ms on normal data.
- Regression policy: fix forward; preserve recent behavior and optimize around it.
- Test strategy: tests after fix plus agent-executed QA.
- Loading policy: show visible loading feedback only when a custom-configuration-related operation exceeds 2 seconds; do not flash under 2 seconds; clear on success and failure.
- Scope default: “custom configuration” means profile selection plus import/export/save/delete-cache/save-profile configuration actions.
- Normal-data default: use existing repository defaults and, if present, a representative workbook from `uploads/` or generated fixture data chosen by the executor; record the chosen fixture in evidence.

### Metis Review (gaps addressed)
- Profile first; do not assume the lag source.
- Distinguish Angular main-thread/render lag, derived-data recomputation, CSS/layout cost, backend/API latency, import/export cost, and recent git changes.
- A delayed loader cannot render while synchronous work blocks the main thread; the implementation must yield/chunk/defer heavy work rather than hiding blocked work behind a spinner.
- Avoid a broad Angular rewrite or new state-management framework.
- Add concrete acceptance criteria for no-loader-under-2s, loader-after-2s, clear-on-success/failure, and derived-view correctness.

## Work Objectives
### Core Objective
Make custom-configuration interactions visibly responsive, with normal operations completing the primary UI update within about 300ms and long operations showing loading feedback after 2 seconds.

### Deliverables
- Timing evidence before and after changes.
- Targeted performance changes to confirmed hotspots.
- Delayed loading UI for long custom-configuration operations.
- Focused automated specs for changed behavior.
- Agent-executed browser/API/build QA evidence.

### Definition of Done (verifiable conditions with commands)
- From `frontend`: `npm run build` exits 0.
- From `frontend`: `npm test -- --watch=false --browsers=ChromeHeadless` exits 0 after adding focused specs.
- From repo root: `.\.venv\Scripts\python.exe app.py` starts Flask at `http://127.0.0.1:5000/`.
- API smoke: `curl http://127.0.0.1:5000/api/config` returns HTTP 200 with JSON.
- Browser QA evidence shows custom-configuration profile selection has immediate visible acknowledgement and primary UI update at or below ~300ms on the selected normal fixture.
- Browser QA evidence shows no loading indicator for a custom-configuration operation completing under 2 seconds.
- Browser QA evidence shows loading indicator appears after at least 2 seconds for an intentionally delayed custom-configuration operation and clears on success/failure.

### Must Have
- Preserve recent unstaged and committed behavior; inspect git status/diff/log before coding.
- Fix forward; do not revert the recent change unless explicitly approved later.
- Keep changes narrow to custom-configuration performance and loading feedback.
- Maintain derived data correctness: prefix verification, price conflict rows, long-code counts, skipped companies, product modal code views, misorder groups, and near-phrase groups.

### Must NOT Have (guardrails, AI slop patterns, scope boundaries)
- MUST NOT rewrite `app.component.ts` wholesale.
- MUST NOT introduce NgRx, signals migration, virtual scrolling libraries, or a new state-management framework.
- MUST NOT optimize unrelated pages/workflows.
- MUST NOT hide main-thread blocking behind a spinner without reducing/yielding heavy work.
- MUST NOT leave loading state stuck after failures.
- MUST NOT require “manual feeling” as proof; all verification must be agent-executed.

## Verification Strategy
> ZERO HUMAN INTERVENTION - all verification is agent-executed.
- Test decision: tests-after using Angular/Karma focused specs; Python tests are not relied on because only an ad hoc `test_logic.py` exists.
- QA policy: Every task has agent-executed happy and failure/edge scenarios.
- Evidence: `.sisyphus/evidence/task-{N}-{slug}.{ext}`

## Execution Strategy
### Parallel Execution Waves
> Parallelism is intentionally limited by the profiling-first requirement: Task 1 must establish evidence before implementation tasks run.
> Wave 2 contains the independent implementation work; single-task Wave 1/3/4 are required by dependencies, not under-splitting.

Wave 1: Task 1 (baseline/git/profile fixture)
Wave 2: Task 2 (derived-data optimization), Task 3 (handler scheduling/yielding), Task 4 (template/render/CSS optimization) — only for hotspots confirmed by Task 1 evidence
Wave 3: Task 5 (delayed loading UI, after Task 3 exposes operation state)
Wave 4: Task 6 (integration verification and commit plan)

### Dependency Matrix (full, all tasks)
| Task | Depends On | Blocks |
| --- | --- | --- |
| 1. Baseline and git context | None | 2, 3, 4, 5, 6 |
| 2. Derived-data recomputation optimization | 1 | 5, 6 |
| 3. Custom-config handler scheduling | 1 | 5, 6 |
| 4. Template/render/CSS optimization | 1 | 6 |
| 5. Delayed loading indicator | 2, 3 | 6 |
| 6. Integration verification and commits | 2, 3, 4, 5 | Final Verification Wave |

### Agent Dispatch Summary (wave → task count → categories)
| Wave | Task Count | Categories |
| --- | ---: | --- |
| 1 | 1 | unspecified-high |
| 2 | 3 | unspecified-high, visual-engineering |
| 3 | 1 | visual-engineering |
| 4 | 1 | unspecified-high |

## TODOs
> Implementation + Test = ONE task. Never separate.
> EVERY task MUST have: Agent Profile + Parallelization + QA Scenarios.

- [ ] 1. Baseline, git context, and reproducible fixture

  **What to do**: Inspect current git status/diff/log read-only, record unstaged files and latest commit summary, choose a representative normal-data fixture, and collect baseline timings for all custom-configuration actions. Use browser Performance API/console timing around `onProfileChange()`, `saveProfileConfig()`, `deleteProfileCache()`, `exportConfig()`, `onConfigImportChange()`, and `importConfig()`. Record whether lag is main-thread, render/layout, backend/API, import/export, or derived-data recomputation.
  **Must NOT do**: Do not stage, commit, revert, or edit source files in this task. Do not assume the hotspot without timing evidence.

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: needs cross-cutting diagnosis and safe git/performance inspection.
  - Skills: [`git-master`] - required for read-only git inspection and later commit style detection.
  - Omitted: [`frontend-ui-ux`] - design changes are not the primary goal in this diagnostic task.

  **Parallelization**: Can Parallel: YES | Wave 1 | Blocks: 2, 3, 4, 5, 6 | Blocked By: none

  **References**:
  - Pattern: `frontend/src/app/app.component.html:12-23` - top custom-configuration controls.
  - Pattern: `frontend/src/app/app.component.html:216` - save configuration control.
  - Handler: `frontend/src/app/app.component.ts:137` - `onProfileChange()`.
  - Handler: `frontend/src/app/app.component.ts:293` - `saveProfileConfig()`.
  - Handler: `frontend/src/app/app.component.ts:306` - `deleteProfileCache()`.
  - Handler: `frontend/src/app/app.component.ts:322` - `exportConfig()`.
  - Handler: `frontend/src/app/app.component.ts:385` - `onConfigImportChange($event)`.
  - Handler: `frontend/src/app/app.component.ts:403` - `importConfig()`.
  - API: `app.py:938` - `GET /api/config` smoke endpoint.

  **Acceptance Criteria**:
  - [ ] Evidence file `.sisyphus/evidence/task-1-baseline.md` lists `git status --short`, latest commit summary, unstaged diff summary, and chosen fixture/data size.
  - [ ] Evidence includes before timings for profile selection, import config, save profile config, delete cache, and export config.
  - [ ] Evidence classifies the dominant lag source(s) and explicitly names any action exceeding 300ms or 2s.
  - [ ] No application files are modified by this task.

  **QA Scenarios**:
  ```
  Scenario: Baseline custom-configuration timing
    Tool: Playwright / browser Performance API
    Steps: Launch Flask with `.\.venv\Scripts\python.exe app.py`; open `http://127.0.0.1:5000/`; select the chosen normal fixture/profile; click each custom-configuration action once; capture timings with performance marks or console timing.
    Expected: `.sisyphus/evidence/task-1-baseline.md` contains concrete ms timings and dominant lag classification.
    Evidence: .sisyphus/evidence/task-1-baseline.md

  Scenario: Git context is read-only
    Tool: Bash
    Steps: Run read-only git status/diff/log commands using git-master environment; do not run add/commit/reset/revert.
    Expected: Evidence records current unstaged and committed context; working tree remains unchanged from before task.
    Evidence: .sisyphus/evidence/task-1-git-context.txt
  ```

  **Commit**: NO | Message: n/a | Files: []

- [ ] 2. Optimize derived-data recomputation with explicit invalidation

  **What to do**: Based on Task 1 timings, optimize only recomputation hotspots explicitly confirmed in Task 1 evidence among `refreshDerivedCodeViews()`, `verifyPrefixes()`, `refreshNearPhraseGroups()`, `buildPriceConflictRows()`, `updateLongCodeCounts()`, and related helper paths. Add memoization/caching only with explicit invalidation triggers: selected profile changes, config import/save/delete, company/product data changes, prefix/word rule changes, and modal selection changes. Convert repeatedly recomputed values used by the template into stored fields only when Task 1 evidence identifies them as repeated render work. Add focused Angular specs for cache invalidation and correctness of changed derivation helpers.
  **Must NOT do**: Do not change output semantics for derived code views. Do not cache across profile/config changes without invalidation. Do not change candidate hotspots that Task 1 documents as “not confirmed.”

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: performance-sensitive TypeScript logic and correctness preservation.
  - Skills: [] - no extra skill required.
  - Omitted: [`frontend-ui-ux`] - logic optimization, not visual redesign.

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: 5, 6 | Blocked By: 1

  **References**:
  - Pattern: `frontend/src/app/app.component.ts:517` - `verifyPrefixes()` currently calls sort/derived refresh.
  - Pattern: `frontend/src/app/app.component.ts:578` - `refreshDerivedCodeViews()` broad recomputation entry point.
  - Pattern: `frontend/src/app/app.component.ts:614` - `updateLongCodeCounts()` loops companies/products and generates codes.
  - Pattern: `frontend/src/app/app.component.ts:760` - `buildPriceConflictRows()` loops companies/products.
  - Pattern: `frontend/src/app/app.component.ts:1143` - `refreshNearPhraseGroups()` nested phrase comparison/edit distance hotspot.
  - Test: `frontend/package.json` - `npm test` exists through Angular/Karma but needs meaningful specs.

  **Acceptance Criteria**:
  - [ ] Every changed derived-data function is named in Task 1 evidence as a confirmed hotspot; every listed candidate that is not changed is documented as “not confirmed” in `.sisyphus/evidence/task-2-scope.md`.
  - [ ] Changed derivation logic has focused `*.spec.ts` coverage for normal recompute and invalidation after profile/config/product data changes.
  - [ ] Derived code views, price conflict rows, long-code counts, skipped companies, modal product codes, misorder groups, and near-phrase groups remain correct on representative data.
  - [ ] Post-fix evidence shows confirmed recomputation hotspots improved versus Task 1 baseline.
  - [ ] `npm test -- --watch=false --browsers=ChromeHeadless` exits 0.

  **QA Scenarios**:
  ```
  Scenario: Derived data remains correct after profile change
    Tool: Bash
    Steps: From `frontend`, run `npm test -- --watch=false --browsers=ChromeHeadless` after adding specs that change profile/config and assert refreshed derived fields.
    Expected: Tests pass and prove cache invalidation updates derived views.
    Evidence: .sisyphus/evidence/task-2-karma.txt

  Scenario: Cache invalidation prevents stale derived data
    Tool: Bash
    Steps: Run a focused spec that changes product/company data after an initial recompute, then asserts long-code counts/price conflicts/near-phrase groups change accordingly.
    Expected: No stale cached values remain after mutation triggers.
    Evidence: .sisyphus/evidence/task-2-invalidation.txt
  ```

  **Commit**: YES | Message: `perf(ui): optimize config derived data recomputation` | Files: [`frontend/src/app/app.component.ts`, `frontend/src/app/*.spec.ts`]

- [ ] 3. Make custom-config handlers yield before heavy work

  **What to do**: Update only custom-configuration handlers that Task 1 evidence names as laggy so clicks provide immediate visual acknowledgement before expensive recomputation. If Task 1 shows synchronous main-thread blocking, split heavy work into awaited micro/macro-task boundaries or chunked work so Angular can render pending/loading state. Apply only to confirmed laggy handlers among `onProfileChange()`, `importConfig()`, `saveProfileConfig()`, `deleteProfileCache()`, and `exportConfig()`; document skipped handlers as “not confirmed.” Keep export synchronous only if Task 1 confirms it is fast.
  **Must NOT do**: Do not introduce arbitrary long timeouts. Do not move business logic to backend unless profiling proves backend is the bottleneck and existing API contract supports it.

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: async UI scheduling and state consistency.
  - Skills: [] - no extra skill required.
  - Omitted: [`git-master`] - no git operation inside this implementation task.

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: 5, 6 | Blocked By: 1

  **References**:
  - Handler: `frontend/src/app/app.component.ts:137` - `onProfileChange()` applies config and triggers broad recomputation.
  - Handler: `frontend/src/app/app.component.ts:403` - `importConfig()` posts to `/api/config/profile/<profile>`.
  - Handler: `frontend/src/app/app.component.ts:293` - `saveProfileConfig()`.
  - Handler: `frontend/src/app/app.component.ts:306` - `deleteProfileCache()`.
  - Backend: `app.py:943` - `POST /api/config`.
  - Backend: `app.py:948` - `POST /api/config/profile/<profile>`.

  **Acceptance Criteria**:
  - [ ] Every changed handler is named in Task 1 evidence as laggy; skipped handlers are documented as “not confirmed” in `.sisyphus/evidence/task-3-scope.md`.
  - [ ] First visible acknowledgement for custom-configuration click occurs before heavy recomputation begins.
  - [ ] Primary UI update for normal fixture completes around 300ms after optimization unless Task 1 proves backend/file work dominates; any exception is documented with timings.
  - [ ] Loading/error state is cleared on success and failure paths.
  - [ ] `npm run build` exits 0.

  **QA Scenarios**:
  ```
  Scenario: Profile selection acknowledges immediately
    Tool: Playwright
    Steps: Start Flask; open `http://127.0.0.1:5000/`; add performance marks around click and first visible acknowledgement; select a concrete available profile.
    Expected: Acknowledgement occurs before heavy derived refresh and primary UI update is around 300ms on the normal fixture.
    Evidence: .sisyphus/evidence/task-3-profile-timing.json

  Scenario: Failed import clears operation state
    Tool: Playwright
    Steps: Intercept `POST /api/config/profile/<profile>` to return HTTP 500; trigger config import with a test file; observe UI state.
    Expected: Operation state clears, error is visible or existing error handling remains intact, and no spinner/loading state remains stuck.
    Evidence: .sisyphus/evidence/task-3-import-failure.png
  ```

  **Commit**: YES | Message: `perf(ui): yield during custom config actions` | Files: [`frontend/src/app/app.component.ts`, `frontend/src/app/*.spec.ts`]

- [ ] 4. Reduce template/render and CSS repaint hotspots

  **What to do**: Replace template-called methods confirmed by Task 1 evidence as high-frequency render work with cached fields updated by explicit refresh methods. Specifically inspect `selectedProfileNote()`, `selectedProfileLabel()`, `wordRuleCount()`, `skippedCompanies()`, `productRowSummary()`, and `productPriceSummary()`; each must be either converted to cached state if confirmed or documented as “not confirmed.” Add `trackBy` functions for the large `*ngFor` loops at company table, product modal, price modal nested lists, misorder list, and near-phrase list when stable identity exists. Gate heavy modal sections so they render only when open and when data exists. If profiling confirms CSS repaint cost, narrow or remove `backdrop-filter`/heavy sticky effects only in large overlays. Add focused specs for any moved computed template values.
  **Must NOT do**: Do not redesign the UI or change Vietnamese labels/text except adding loading copy if needed. Do not remove useful information from modals.

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: template/render/CSS performance with visual regression awareness.
  - Skills: [`frontend-ui-ux`] - needed for preserving UI clarity while reducing repaint cost.
  - Omitted: [`git-master`] - no git operation inside this implementation task.

  **Parallelization**: Can Parallel: YES | Wave 2 | Blocks: 6 | Blocked By: 1

  **References**:
  - Template: `frontend/src/app/app.component.html:26` - profile note/label method usage.
  - Template: `frontend/src/app/app.component.html:129`, `:215`, `:331`, `:332` - repeated skipped-company derived usage.
  - Template: `frontend/src/app/app.component.html:161` - large company table `*ngFor`.
  - Template: `frontend/src/app/app.component.html:231` - product modal rows.
  - Template: `frontend/src/app/app.component.html:354`, `:370`, `:394` - price modal nested loops.
  - Template: `frontend/src/app/app.component.html:424`, `:466` - misorder/near-phrase loops.
  - CSS: `frontend/src/app/app.component.css:64`, `:620`, `:721` - `backdrop-filter` repaint candidates.
  - CSS: `frontend/src/app/app.component.css:334`, `:366`, `:611` - sticky regions/actions.

  **Acceptance Criteria**:
  - [ ] `.sisyphus/evidence/task-4-scope.md` lists `selectedProfileNote()`, `selectedProfileLabel()`, `wordRuleCount()`, `skippedCompanies()`, `productRowSummary()`, and `productPriceSummary()` as either converted to cached state or “not confirmed by Task 1.”
  - [ ] Company table, product modal, price modal nested lists, misorder list, and near-phrase list either use explicit `trackBy` functions or are documented in `.sisyphus/evidence/task-4-trackby.md` as lacking stable identity.
  - [ ] Playwright screenshots for company table, product modal, price modal, misorder section, and near-phrase section show labels/rows present, and browser console has zero errors.
  - [ ] `npm run build` exits 0.

  **QA Scenarios**:
  ```
  Scenario: Template render remains stable on large lists
    Tool: Playwright
    Steps: Load representative fixture; open company table, product modal, price modal, misorder/near-phrase sections; capture screenshot after each view.
    Expected: Content appears with no missing rows/sections and no console errors.
    Evidence: .sisyphus/evidence/task-4-render-screenshots.zip

  Scenario: Cached template values update after data changes
    Tool: Bash
    Steps: Run focused Angular specs for cached skipped-company/profile-label/summary values after config/data mutation.
    Expected: Cached fields refresh exactly when source data changes.
    Evidence: .sisyphus/evidence/task-4-cache-specs.txt
  ```

  **Commit**: YES | Message: `perf(ui): reduce custom config render work` | Files: [`frontend/src/app/app.component.ts`, `frontend/src/app/app.component.html`, `frontend/src/app/app.component.css`, `frontend/src/app/*.spec.ts`]

- [ ] 5. Add delayed loading feedback for long custom-config operations

  **What to do**: Add operation-level loading state for custom-configuration actions. Start a 2-second timer when an operation begins; show loading UI only if the operation is still active after the timer fires; clear timer and indicator on success, failure, cancellation, and superseded operations. Use accessible text and existing visual language. Suggested copy: `Đang tải cấu hình...` for profile/import/save/delete config operations, with action-specific text only if the existing UI needs it. Add fakeAsync/tick specs for under-2s/no-flash, after-2s/show, success clear, and failure clear.
  **Must NOT do**: Do not show the loading indicator immediately. Do not let a delayed timer from an old operation show after a newer operation starts.

  **Recommended Agent Profile**:
  - Category: `visual-engineering` - Reason: visible UI state plus accessible styling.
  - Skills: [`frontend-ui-ux`] - preserve existing UI while adding unobtrusive feedback.
  - Omitted: [`git-master`] - no git operation inside this implementation task.

  **Parallelization**: Can Parallel: YES | Wave 3 | Blocks: 6 | Blocked By: 2, 3

  **References**:
  - Template: `frontend/src/app/app.component.html:12-23` - top custom configuration actions.
  - Template: `frontend/src/app/app.component.html:216` - save configuration button area.
  - Logic: `frontend/src/app/app.component.ts:137`, `:293`, `:306`, `:385`, `:403` - operations requiring delayed loading coverage.
  - Styles: `frontend/src/app/app.component.css` - existing component styling to follow.

  **Acceptance Criteria**:
  - [ ] Loading indicator does not appear for operations completing before 2 seconds.
  - [ ] Loading indicator appears at/after 2 seconds for still-running custom-config operation.
  - [ ] Loading indicator clears after operation success.
  - [ ] Loading indicator clears after operation failure.
  - [ ] Superseded operations cannot leak stale loading indicators.
  - [ ] `npm test -- --watch=false --browsers=ChromeHeadless` exits 0.

  **QA Scenarios**:
  ```
  Scenario: Fast operation has no loading flash
    Tool: Bash
    Steps: Run Angular fakeAsync spec that starts a custom-config operation, completes it before 2000ms, then advances timers past 2000ms.
    Expected: Loading-visible state remains false and no stale timer changes it afterward.
    Evidence: .sisyphus/evidence/task-5-fast-no-flash.txt

  Scenario: Slow operation shows and clears loading
    Tool: Playwright
    Steps: Intercept `POST /api/config/profile/<profile>` or selected config API with >2200ms delay; trigger custom config action; capture at 1500ms, 2100ms, and after completion.
    Expected: No indicator at 1500ms; indicator visible at 2100ms; indicator gone after completion.
    Evidence: .sisyphus/evidence/task-5-delayed-loading.zip

  Scenario: Superseded slow operation cannot leak loading
    Tool: Bash
    Steps: Run Angular fakeAsync spec that starts operation A, advances to 1900ms, starts operation B, completes operation A, advances past A's old 2000ms boundary, then completes operation B before/after its own timer as separate assertions.
    Expected: Operation A's timer never shows stale loading for operation B; final loading state is false after B completes.
    Evidence: .sisyphus/evidence/task-5-superseded.txt
  ```

  **Commit**: YES | Message: `feat(ui): show delayed config loading state` | Files: [`frontend/src/app/app.component.ts`, `frontend/src/app/app.component.html`, `frontend/src/app/app.component.css`, `frontend/src/app/*.spec.ts`]

- [ ] 6. Integration verification, build, and atomic commit finalization

  **What to do**: Run full agreed verification, compare before/after timings against Task 1, capture final evidence, and create atomic commit(s) following detected repository style. Commit groups must pair tests with implementation and split by concern if more than two files changed across logic/render/loading. If the repository has only one initial commit style, use concise semantic messages shown in this plan unless git log indicates otherwise.
  **Must NOT do**: Do not commit generated build artifacts unless they were already tracked and intentionally changed by the project workflow. Do not include `.sisyphus/evidence` unless the repository convention requires it.

  **Recommended Agent Profile**:
  - Category: `unspecified-high` - Reason: multi-command verification and commit discipline.
  - Skills: [`git-master`] - required for commit style detection, staging, and atomic commits.
  - Omitted: [`frontend-ui-ux`] - visual decisions should already be complete.

  **Parallelization**: Can Parallel: NO | Wave 4 | Blocks: Final Verification Wave | Blocked By: 2, 3, 4, 5

  **References**:
  - Build: `frontend/package.json` - `npm run build`.
  - Tests: `frontend/package.json` - `npm test` / `ng test`.
  - App run: `app.py` `main()` runs Flask on `127.0.0.1:5000`.
  - API: `app.py:938` - `GET /api/config`.
  - Commit policy: use git-master skill; inspect `git status`, `git diff`, and `git log --oneline -10` before committing.

  **Acceptance Criteria**:
  - [ ] `npm run build` from `frontend` exits 0.
  - [ ] `npm test -- --watch=false --browsers=ChromeHeadless` from `frontend` exits 0.
  - [ ] Flask starts with `.\.venv\Scripts\python.exe app.py` and serves `http://127.0.0.1:5000/`.
  - [ ] `curl http://127.0.0.1:5000/api/config` returns HTTP 200 JSON.
  - [ ] Final timing evidence shows primary custom-config update near 300ms on normal fixture or documents an unavoidable external bottleneck with delayed loading working.
  - [ ] Atomic commits are created according to git-master rules, with implementation and tests paired.

  **QA Scenarios**:
  ```
  Scenario: Full local verification passes
    Tool: Bash
    Steps: From `frontend`, run `npm run build`; run `npm test -- --watch=false --browsers=ChromeHeadless`; from repo root start Flask; run API smoke for `/api/config`.
    Expected: All commands exit 0/HTTP 200 and outputs are saved.
    Evidence: .sisyphus/evidence/task-6-verification.txt

  Scenario: Browser regression for custom config
    Tool: Playwright
    Steps: Open Flask-served app; execute profile selection, import config failure path, save profile config, delete cache confirmation/cancellation path if present, and export config; capture console errors and timing marks.
    Expected: No console errors, no stuck loading state, normal profile update around 300ms, delayed loader behavior only after 2 seconds.
    Evidence: .sisyphus/evidence/task-6-browser-regression.zip
  ```

  **Commit**: YES | Message: `perf(ui): improve custom config responsiveness` / split by git-master atomic plan | Files: [changed source/test files only]

## Final Verification Wave (MANDATORY — after ALL implementation tasks)
> 4 review agents run in PARALLEL. ALL must APPROVE using agent-executed evidence. Present consolidated results to user afterward.
> User acknowledgement is a post-verification completion checkpoint only; it is not an acceptance criterion and must not replace automated/build/API/browser QA evidence.
> **Do NOT mark F1-F4 as checked until all four agents approve and the consolidated report is presented.** Rejection or user feedback -> fix -> re-run -> present again.
- [ ] F1. Plan Compliance Audit — oracle
- [ ] F2. Code Quality Review — unspecified-high
- [ ] F3. Real Manual QA — unspecified-high (+ playwright for UI)
- [ ] F4. Scope Fidelity Check — deep

## Commit Strategy
- Before any commit: inspect `git status`, `git diff`, and `git log --oneline -10` using git-master.
- Preserve unstaged user changes unless they are the files intentionally changed for this task; never stage unrelated files.
- Split commits by concern if multiple concerns changed:
  1. Derived recomputation performance + direct specs.
  2. Handler scheduling/render-yield performance + direct specs.
  3. Template/CSS render performance + direct specs.
  4. Delayed loading UI + direct specs.
- If implementation touches only one cohesive area, git-master may combine implementation and its direct specs, but must justify grouping per skill rules.
- Do not commit generated `static/`, `dist/`, `build/`, or `deploy/` outputs unless they are already tracked and intentionally part of the project’s release workflow.

## Success Criteria
- Custom-configuration interactions no longer visibly freeze on the selected normal fixture.
- Primary custom-config UI update completes around 300ms on normal data, or timing evidence identifies a non-UI external bottleneck and delayed feedback still works.
- Loading feedback appears only after 2 seconds for still-running custom-config operations and clears reliably.
- Existing profile/config/import/export/save/delete cache behavior remains functionally equivalent.
- Build, focused tests, API smoke, and browser QA all pass with saved evidence.
