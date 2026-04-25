---
phase: 07-ci-and-repo-hygiene
plan: 02
subsystem: ci-and-testing
tags: [ci, github-actions, pytest, platform-guard, stale-name, version-check]

# Dependency graph
requires: [07-01]
provides:
  - stale-name CI job passes without false-positives on its own command string
  - version-check CI job enforces config.py vs package.json version consistency
  - test_restore_maximized_window skips on Linux CI (no ImportError)
  - test_app_config.py docstring free of banned "Windows Toolbox" literal
affects: [07-03, 07-04]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "@pytest.mark.skipif(sys.platform != 'win32') for Win32-only test functions"
    - "grep -oP Perl-mode regex for version extraction in CI shell step"

key-files:
  created: []
  modified:
    - .github/workflows/ci.yml
    - tests/unit/test_app_config.py
    - tests/unit/test_snap_geometry.py

key-decisions:
  - "Add --exclude-dir=.github to stale-name grep so ci.yml does not self-match (D-05)"
  - "Add version-check job with grep -oP extraction and fail-on-mismatch logic (D-10)"
  - "Rephrase test_app_config.py docstring to use 'not the old product name' instead of literal banned string (D-06)"
  - "Add @pytest.mark.skipif(sys.platform != 'win32') to test_restore_maximized_window to prevent ctypes.wintypes ImportError on Linux CI (D-07)"

requirements-completed: [CI-03, CI-04]

# Metrics
duration: 1min
completed: 2026-04-25
---

# Phase 7 Plan 02: CI False-Positive Fixes and Platform Guard Summary

**Fixed two CI false-positive sources and added version-check job: stale-name grep no longer self-matches ci.yml, win32 test skips on Linux, and version drift is now caught by CI.**

## Performance

- **Duration:** 1min
- **Started:** 2026-04-25T01:46:36Z
- **Completed:** 2026-04-25T01:47:41Z
- **Tasks:** 2/2
- **Files modified:** 3

## Accomplishments

- Added `--exclude-dir=.github` to the stale-name grep in ci.yml so the workflow file no longer matches its own search string
- Added `version-check` job to ci.yml that extracts APP_VERSION from config.py and package.json via `grep -oP` and fails if they differ
- Rephrased the `test_app_name_is_virelo` docstring in test_app_config.py to use "not the old product name" instead of the literal "Windows Toolbox" string
- Added `import sys` and `import pytest` at module level in test_snap_geometry.py
- Added `@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs only available on Windows")` to `test_restore_maximized_window` to prevent `ImportError: cannot import name 'wintypes' from 'ctypes'` on Ubuntu CI

## Task Commits

Each task was committed atomically:

1. **Task 1: Fix stale-name grep false positive and add version-check job to ci.yml** - `dbc3bee` (fix)
2. **Task 2: Fix test_app_config.py docstring and add platform guard to test_snap_geometry.py** - `8fd54d7` (fix)

## Files Created/Modified

- `.github/workflows/ci.yml` - Added `--exclude-dir=.github` to stale-name grep; added `version-check` job (5th job) with grep-oP version extraction and fail-on-mismatch logic
- `tests/unit/test_app_config.py` - Rephrased docstring on `test_app_name_is_virelo` to remove literal banned string
- `tests/unit/test_snap_geometry.py` - Added `import sys` and `import pytest` at module level; added `@pytest.mark.skipif` decorator on `test_restore_maximized_window`

## Decisions Made

- `--exclude-dir=.github` (not `--exclude=.github/workflows/ci.yml`) used for directory-level exclusion, consistent with existing `--exclude-dir` flags in the same command
- `grep -oP` Perl-mode regex chosen for version extraction per D-10 and research Pattern 3; no new tooling needed on ubuntu-latest
- Import order in test_snap_geometry.py follows isort convention: `import sys` before `import pytest` (both stdlib/third-party), then existing `from unittest.mock` import

## Deviations from Plan

None - plan executed exactly as written.

## Self-Check: PASSED

- `grep "exclude-dir=.github" .github/workflows/ci.yml` exits 0 -- line 61 confirmed
- `grep "version-check:" .github/workflows/ci.yml` exits 0 -- line 67 confirmed
- `grep -c "runs-on: ubuntu-latest" .github/workflows/ci.yml` returns 5 -- confirmed
- `grep "Windows Toolbox" tests/unit/test_app_config.py` exits 1 (no match) -- confirmed
- `grep "not the old product name" tests/unit/test_app_config.py` exits 0 -- line 49 confirmed
- `grep "pytest.mark.skipif" tests/unit/test_snap_geometry.py` exits 0 -- line 111 confirmed
- `grep "^import sys" tests/unit/test_snap_geometry.py` exits 0 -- line 7 confirmed
- `grep "^import pytest" tests/unit/test_snap_geometry.py` exits 0 -- line 9 confirmed
- `pytest tests/unit/ -q` exits 0 -- 66 passed in 0.04s
- Commits `dbc3bee` and `8fd54d7` verified in git log

---
*Phase: 07-ci-and-repo-hygiene*
*Completed: 2026-04-25*
