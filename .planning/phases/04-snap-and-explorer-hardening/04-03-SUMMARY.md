---
phase: 04-snap-and-explorer-hardening
plan: 03
subsystem: snap-geometry-tests
tags: [testing, snap, multi-monitor, geometry, restore]
dependency_graph:
  requires: [04-01]
  provides: [SNAP-05-coverage, SNAP-04-coverage]
  affects: []
tech_stack:
  added: []
  patterns: [ShiftSnapRestore.__new__ bypass for unit testing, create=True patch for stub modules]
key_files:
  created: []
  modified:
    - tests/unit/test_snap_geometry.py
decisions:
  - Used _make_settings() helper instead of conftest MockSettings import (conftest not directly importable from unit test subdir)
  - Used patch create=True for PySide6.QtWidgets.QApplication since conftest stub module lacks that attribute
metrics:
  duration: 3min
  completed: "2026-04-24T21:44:09Z"
---

# Phase 4 Plan 3: Multi-Monitor Snap Geometry Tests Summary

Three new unit tests covering negative-coordinate monitor geometry, vertical multi-monitor layout, and maximized-window restore round-trip via SW_MAXIMIZE.

## What Was Done

### Task 1: Negative-coordinates and vertical-layout snap geometry tests
**Commit:** `2dfaa06`

Added two pure geometry tests to `tests/unit/test_snap_geometry.py`:

- `test_calculate_snap_position_negative_coords` -- validates that `calculate_snap_position` correctly computes x offset when the monitor has a negative origin (left-of-primary layout, monitor_left=-1920). Asserts x = -1690, confirming the monitor_left offset is applied.
- `test_calculate_snap_position_vertical_layout` -- validates y offset when the monitor is below the primary (monitor_top=1080, size 2560x1440). Asserts y = 1253, confirming monitor_top is applied to vertical centering.

### Task 2: Maximized-restore round-trip test
**Commit:** `0b25a16`

Added `test_restore_maximized_window` validating SNAP-04: when a previously-maximized window is restored, `ShiftSnapRestore._restore` calls `win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)` instead of repositioning to the old geometry.

The test uses `ShiftSnapRestore.__new__` to bypass `__init__` (avoids keyboard hooks and EnumWindows in test context), patches `get_monitor_rect`, `USER32.GetWindowRect`, and `QApplication.topLevelWidgets` to isolate the restore logic.

Added `_make_settings()` helper and `unittest.mock` imports to support the restore test.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] conftest.MockSettings not importable from unit test subdir**
- **Found during:** Task 2
- **Issue:** `from conftest import MockSettings` fails with `ModuleNotFoundError` because conftest.py is auto-loaded by pytest but not on the Python path for direct imports from `tests/unit/`.
- **Fix:** Created `_make_settings()` helper function inline in the test file that replicates MockSettings behavior using `DEFAULTS` from `virelo.app.config`.
- **Files modified:** tests/unit/test_snap_geometry.py
- **Commit:** 0b25a16

**2. [Rule 3 - Blocking] PySide6.QtWidgets stub lacks QApplication attribute**
- **Found during:** Task 2
- **Issue:** `patch("PySide6.QtWidgets.QApplication", ...)` raises `AttributeError` because the conftest stub module for PySide6.QtWidgets does not define QApplication.
- **Fix:** Used `create=True` parameter on the patch call to allow patching a non-existent attribute on the stub module.
- **Files modified:** tests/unit/test_snap_geometry.py
- **Commit:** 0b25a16

## Verification

- `pytest tests/unit/test_snap_geometry.py -v` -- 11/11 tests pass (8 original + 3 new)
- `grep -c "def test_" tests/unit/test_snap_geometry.py` -- returns 11
- `pytest tests/unit/ -v` -- 53/53 full unit suite passes

## Self-Check: PASSED
