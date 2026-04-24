---
phase: 03-structure-and-quality
plan: 01
subsystem: quality-infrastructure
tags: [ruff, pytest, vitest, pyproject, linting, formatting, test-infra]
dependency_graph:
  requires: []
  provides: [pyproject-toml, ruff-baseline, pytest-config, vitest-config]
  affects: [all-python-files, frontend-config]
tech_stack:
  added: [ruff-0.15.12, vitest-4.1.5, testing-library-react, testing-library-jest-dom, jsdom]
  patterns: [implicit-string-concatenation-for-long-log-lines]
key_files:
  created:
    - pyproject.toml
    - frontend/src/test-setup.js
  modified:
    - app_config.py
    - bridge.py
    - main.py
    - settings_state.py
    - webview.py
    - explorer_columns.py
    - capture_guard.py
    - settings.py
    - snap_service.py
    - startup_shortcut.py
    - theme.py
    - workers.py
    - scripts/build-icon.py
    - scripts/build-installer-bmps.py
    - frontend/vite.config.js
    - frontend/package.json
    - frontend/package-lock.json
decisions:
  - "Ruff E/F/I/UP rules with line-length 100 and double-quote format established as baseline"
  - "Unused variable assignments removed rather than suppressed where safe"
  - "Long log format strings broken via implicit concatenation rather than per-file-ignores"
metrics:
  duration: 6min
  completed: 2026-04-24
---

# Phase 03 Plan 01: Quality Tooling Setup Summary

Ruff lint/format config and Vitest test infrastructure established with zero-violation baseline across all Python files

## What Was Done

### Task 1: pyproject.toml and Ruff Baseline
- Created `pyproject.toml` with complete project metadata, dependency groups (runtime, dev, build), Ruff configuration (E/F/I/UP rules, line-length 100, py312 target), and pytest configuration (testpaths, markers)
- Installed ruff 0.15.12 and ran `ruff check --fix .` to auto-fix 101 issues: import sorting (I001), pyupgrade modernizations (UP006/UP007/UP035/UP045), unused imports (F401)
- Ran `ruff format .` to apply consistent formatting (double quotes, line-length 100) across 11 files
- Manually fixed 18 remaining E501 line-too-long violations using implicit string concatenation for log format strings
- Removed unused variable assignments (F841) in main.py where `win_left/win_top/win_right/win_bottom` were extracted from `rc` but never referenced

### Task 2: Vitest Test Infrastructure
- Installed vitest@4.1.5, @testing-library/react@16.3.2, @testing-library/jest-dom@6.9.1, @testing-library/user-event@14.6.1, jsdom@29.0.2 as devDependencies
- Added `test` configuration block to `vite.config.js` with jsdom environment, globals, and setupFiles
- Created `frontend/src/test-setup.js` importing `@testing-library/jest-dom` matchers
- Added `"test": "vitest run"` script to `package.json`
- Verified Vitest runs cleanly (discovers zero tests, exits code 0 with --passWithNoTests)

## Verification Results

| Check | Result |
|-------|--------|
| `ruff check .` | All checks passed (exit 0) |
| `ruff format --check .` | 14 files already formatted (exit 0) |
| `npx vitest run --passWithNoTests` | No test files found, exit 0 |
| pyproject.toml contains [tool.ruff] | Confirmed |
| pyproject.toml contains [tool.pytest.ini_options] | Confirmed |
| vite.config.js contains test block | Confirmed |
| test-setup.js contains jest-dom | Confirmed |

## Deviations from Plan

None - plan executed exactly as written.

## Commits

| Task | Commit | Description |
|------|--------|-------------|
| 1 | f03cb86 | Create pyproject.toml and establish Ruff lint/format baseline |
| 2 | 0a8efee | Configure Vitest test infrastructure with jsdom and testing-library |

## Self-Check: PASSED

All files exist, all commits found, all content checks verified.
