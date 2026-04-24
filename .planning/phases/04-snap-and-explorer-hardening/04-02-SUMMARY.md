---
phase: 04-snap-and-explorer-hardening
plan: 02
subsystem: services
tags: [pyside6, qthread, explorer, lifecycle, facade]

# Dependency graph
requires:
  - phase: 04-01
    provides: HotkeyListener class in snap.py, SnapService.set_listener method
provides:
  - ExplorerService module for explorer worker lifecycle management
  - MainWindow wired with HotkeyListener and ExplorerService
  - Bridge side-effects delegating to _hotkey_listener
affects: [04-03]

# Tech tracking
tech-stack:
  added: []
  patterns: [service-facade extraction from MainWindow]

key-files:
  created:
    - virelo/services/explorer_service.py
  modified:
    - virelo/app/window.py
    - virelo/bridge/bridge.py

key-decisions:
  - "ExplorerService follows SnapService facade pattern for consistency"
  - "Thin shim _update_explorer_autosize_thread kept to avoid breaking bridge.py callsite"
  - "Unused time import removed from window.py after extraction"

patterns-established:
  - "Service facade: lifecycle management extracted from MainWindow into dedicated service classes"
  - "Signal delegation: HotkeyListener owns keyboard hooks, bridge delegates key changes to _hotkey_listener"

requirements-completed: [EXPL-01, EXPL-02, EXPL-03]

# Metrics
duration: 4min
completed: 2026-04-24
---

# Phase 04 Plan 02: ExplorerService Extraction and HotkeyListener Wiring Summary

**ExplorerService facade extracts explorer worker lifecycle from MainWindow, HotkeyListener wired for keyboard hook management, bridge delegates key binding changes to _hotkey_listener**

## Performance

- **Duration:** 4min
- **Started:** 2026-04-24T21:40:39Z
- **Completed:** 2026-04-24T21:44:58Z
- **Tasks:** 2
- **Files modified:** 3

## Accomplishments
- Created ExplorerService module with start/stop/is_running lifecycle following SnapService facade pattern
- Wired HotkeyListener into MainWindow signal graph (triggered -> ShiftSnapRestore.perform)
- Removed all inline explorer thread management from MainWindow (110+ lines removed)
- Updated all bridge side-effect delegation from shift_mgr to _hotkey_listener

## Task Commits

Each task was committed atomically:

1. **Task 1: Create ExplorerService module** - `f162115` (feat)
2. **Task 2: Wire HotkeyListener and ExplorerService into MainWindow and update bridge** - `d1ae9cd` (feat)

## Files Created/Modified
- `virelo/services/explorer_service.py` - New ExplorerService class with start/stop/is_running lifecycle, wrapper functions moved from window.py
- `virelo/app/window.py` - Imports HotkeyListener and ExplorerService, delegates lifecycle, removes inline thread management
- `virelo/bridge/bridge.py` - Side effects delegate key binding changes to _hotkey_listener instead of shift_mgr

## Decisions Made
- ExplorerService follows SnapService facade pattern (plain class, not QObject, wraps QThread + worker)
- Kept _update_explorer_autosize_thread as thin shim delegating to _explorer_service.start() to avoid breaking bridge.py callsite
- Removed unused `time` import from window.py after extracting time.sleep to ExplorerService

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Fixed additional shift_mgr reference in bridge.py reset_defaults**
- **Found during:** Task 2 (bridge.py modifications)
- **Issue:** bridge.py line 136-139 (reset_defaults) also referenced shift_mgr.update_binding/update_restore_key/update_press_limit, not covered in plan's line references (only 260-283 were specified)
- **Fix:** Updated reset_defaults to use _hotkey_listener references, same as _apply_side_effects
- **Files modified:** virelo/bridge/bridge.py
- **Verification:** grep confirms zero shift_mgr references remain in bridge.py
- **Committed in:** d1ae9cd (Task 2 commit)

**2. [Rule 1 - Bug] Removed unused time import from window.py**
- **Found during:** Task 2 (cleanup after extraction)
- **Issue:** `import time` was no longer used after time.sleep(0.05) moved to ExplorerService.stop()
- **Fix:** Removed the import line
- **Files modified:** virelo/app/window.py
- **Verification:** grep confirms no time. references remain in window.py
- **Committed in:** d1ae9cd (Task 2 commit)

---

**Total deviations:** 2 auto-fixed (2 bug fixes)
**Impact on plan:** Both auto-fixes necessary for correctness and clean code. No scope creep.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- ExplorerService and HotkeyListener fully wired into MainWindow
- All unit tests pass (11/11)
- Ready for plan 04-03 tests (running in parallel)

---
## Self-Check: PASSED

All files exist, all commits found in git log.

---
*Phase: 04-snap-and-explorer-hardening*
*Completed: 2026-04-24*
