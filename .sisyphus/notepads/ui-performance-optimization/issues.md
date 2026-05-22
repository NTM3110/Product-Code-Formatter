# Issues

## 2026-05-20 Task 1 Git Context

- Git read-only evidence was blocked because the shell wrapper prepended POSIX-style `set ... &&` environment setup that PowerShell rejects before `git` runs. No git mutation commands were run.
