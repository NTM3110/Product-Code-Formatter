# Task 5 Delayed Loading Evidence

Automated spec coverage:
- Fast operation completes before 2 seconds without showing loading.
- Slow operation shows loading after 2 seconds and clears on completion.
- Superseded operation timers cannot leak stale loading into the new operation.

Playwright delayed API scenario:
- Intercepted `POST /api/config` with a 2300 ms delay.
- At 1500 ms: `.config-loading` count was 0.
- At 2250 ms: `.config-loading` count was 1.
- After completion: `.config-loading` count was 0.
