---
phase: "08"
plan: "02"
subsystem: "entry-point, build-pipeline"
tags: [smoke-test, release-verification, CLI, PowerShell]
dependency_graph:
  requires: []
  provides: ["--smoke-test flag", "expanded verify-release.ps1"]
  affects: ["virelo/app/__main__.py", "scripts/verify-release.ps1"]
tech_stack:
  added: ["argparse (stdlib)"]
  patterns: ["non-interactive smoke test", "version cross-check", "bundled asset validation"]
key_files:
  created: []
  modified:
    - "virelo/app/__main__.py"
    - "scripts/verify-release.ps1"
decisions:
  - "--smoke-test parsed before admin elevation to avoid UAC loop (Pitfall 3)"
  - "add_help=False and parse_known_args to avoid interfering with Qt arguments"
  - "All 4 new verify-release checks use PowerShell cmdlets (no $LASTEXITCODE concern)"
metrics:
  duration: "2min"
  completed: "2026-04-25T02:52:54Z"
---

# Phase 8 Plan 2: Smoke Test and Release Verification Summary

Non-interactive --smoke-test flag verifying 6 subsystem checks (icon, frontend, QWebEngine, Settings, SettingsState, VireloBridge) plus 4 new verify-release.ps1 checks for version cross-check, bundled assets, and stale naming detection in dist/.

## Commits

| Task | Commit | Message |
|------|--------|---------|
| 1 | 4e65851 | feat(08-02): add --smoke-test flag and smoke test runner to __main__.py |
| 2 | bf1f30e | feat(08-02): expand verify-release.ps1 with version cross-check and asset validation |

## Task Details

### Task 1: Add --smoke-test flag and smoke test runner to __main__.py

Added `import argparse` and argument parsing in `main()` before admin elevation. The `_run_smoke_test()` function creates a QApplication, runs 6 subsystem checks (icon.ico resource path, frontend/dist/index.html, QWebEngine construction, Settings read/write, SettingsState defaults, VireloBridge initialization), and returns 0/1 based on pass/fail. No window is shown and no UAC prompt is triggered.

Key design choices:
- `add_help=False` prevents argparse from consuming Qt's -h flag
- `parse_known_args` ignores Qt platform arguments like -platform
- Smoke test path exits via `sys.exit()` before any admin elevation, mutex, or MainWindow

### Task 2: Expand verify-release.ps1 with version cross-check and bundled asset validation

Added 4 new check blocks after existing checks and before the Report section:
1. Version cross-check: config.py APP_VERSION vs frontend/package.json version
2. Bundled icon.ico: verifies dist/Virelo/icon.ico exists
3. Bundled frontend: verifies dist/Virelo/frontend/dist/index.html exists
4. Stale naming: scans dist/Virelo/ recursively for files matching "toolbox" pattern

All checks use PowerShell cmdlets (Get-Content, ConvertFrom-Json, Test-Path, Get-ChildItem, Where-Object) and accumulate failures into the existing $errors pattern.

## Deviations from Plan

None - plan executed exactly as written.

## Verification Results

- `_run_smoke_test` function exists (AST parse confirmed)
- `--smoke-test` flag parsed before `_is_admin()` call (line 181 vs line 208)
- `argparse` and `parse_known_args` present in source
- All 4 new verify-release.ps1 checks present (pkgJsonVersion, dist/Virelo/icon.ico, dist/Virelo/frontend/dist/index.html, toolbox scan)
- Existing checks unchanged
- 88 unit tests pass
- Ruff linting clean

## Self-Check: PASSED

- [x] virelo/app/__main__.py exists and contains _run_smoke_test
- [x] scripts/verify-release.ps1 exists and contains all 4 new checks
- [x] Commit 4e65851 verified in git log
- [x] Commit bf1f30e verified in git log
