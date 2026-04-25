---
phase: "08"
plan: "01"
subsystem: "window-chrome"
tags: [nativeEvent, HTCAPTION, signed-lParam, UPX, multi-monitor]
dependency_graph:
  requires: []
  provides: [HTCAPTION-drag-zone, signed-coordinate-decoding, UPX-disabled]
  affects: [virelo/app/window.py, Virelo.spec]
tech_stack:
  added: []
  patterns: [ctypes.c_short-signed-extraction, HTCAPTION-drag-zone-exclusion]
key_files:
  created:
    - tests/unit/test_nchittest.py
  modified:
    - virelo/app/window.py
    - Virelo.spec
decisions:
  - "TITLE_BAR_HEIGHT=35 matches frontend TitleBar (34px+1px border)"
  - "CONTROLS_WIDTH=60 excludes two 28px buttons plus safety margin from drag zone"
  - "pos.x() >= BORDER condition ensures left resize border takes priority over drag"
metrics:
  duration: "2min"
  completed: "2026-04-25"
---

# Phase 8 Plan 1: Window Chrome Fixes Summary

Fixed frameless window dragging via HTCAPTION return in nativeEvent, corrected signed 16-bit lParam extraction for multi-monitor support, and disabled UPX compression in Virelo.spec.

## Commits

| Task | Name | Commit | Files |
|------|------|--------|-------|
| 1 | Unit tests for signed lParam and HTCAPTION hit zones | 54553b6 | tests/unit/test_nchittest.py |
| 2 | Fix nativeEvent with signed lParam and HTCAPTION drag zone | f53251e | virelo/app/window.py |
| 3 | Disable UPX in Virelo.spec | 62ff70f | Virelo.spec |

## What Changed

### Task 1: Unit tests (22 tests)
Created `tests/unit/test_nchittest.py` with pure-logic unit tests covering both signed extraction and hit-zone classification. Six tests verify `ctypes.c_short` decoding of unsigned lParam values to signed 16-bit coordinates (including negative monitor coords like -1920). Sixteen tests verify the hit-zone classification logic: HTCAPTION for the title bar drag zone, all eight edge/corner resize zones, controls area exclusion, and boundary conditions.

### Task 2: nativeEvent signed lParam + HTCAPTION
Modified `virelo/app/window.py` with three changes:
1. Added `TITLE_BAR_HEIGHT = 35` and `CONTROLS_WIDTH = 60` module-level constants
2. Replaced unsigned lParam extraction (`msg.lParam & 0xFFFF`) with signed extraction (`ctypes.c_short(msg.lParam & 0xFFFF).value`) for both x and y coordinates
3. Added HTCAPTION return after the edge/corner block: returns hit code 2 when the mouse is in the top 35px of the window, excluding the 4px resize border and the rightmost 60px where window control buttons live

### Task 3: Disable UPX
Changed `upx=True` to `upx=False` in both the EXE block (line 67) and COLLECT block (line 81) of `Virelo.spec`. This prevents intermittent PySide6/Qt startup crashes caused by UPX compression.

## Deviations from Plan

None -- plan executed exactly as written.

## Verification Results

- 22/22 unit tests pass in `tests/unit/test_nchittest.py`
- 88/88 total unit tests pass (full suite, no regressions)
- `upx=True` no longer appears in `Virelo.spec`
- `upx=False` appears exactly 2 times in `Virelo.spec`
- `TITLE_BAR_HEIGHT = 35` present in `virelo/app/window.py`
- `CONTROLS_WIDTH = 60` present in `virelo/app/window.py`
- Unsigned extraction lines removed from `nativeEvent`

## Threat Surface Scan

No new threat surface introduced. The HTCAPTION drag zone is constrained by the existing edge/corner priority checks and the CONTROLS_WIDTH exclusion, matching threat model T-08-02 mitigation.

## Self-Check: PASSED
