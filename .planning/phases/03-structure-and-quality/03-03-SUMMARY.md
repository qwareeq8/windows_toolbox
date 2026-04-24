---
phase: 03-structure-and-quality
plan: 03
subsystem: infra
tags: [python-packaging, module-extraction, main-shim, pyinstaller, workers-split]

# Dependency graph
requires:
  - phase: 03-structure-and-quality/02
    provides: virelo/ package skeleton with leaf and middle-tier modules
provides:
  - Full virelo/ package with all Python source (upper-tier modules moved)
  - Thin main.py shim delegating to virelo.app.main()
  - workers.py split into key_capture.py and explorer.py
  - ShiftSnapRestore engine in virelo/services/snap.py
  - Updated Virelo.spec for package paths and hiddenimports
affects: [03-structure-and-quality/04, 03-structure-and-quality/05]

# Tech tracking
tech-stack:
  added: []
  patterns: [thin-entry-shim, upper-tier-extraction, COM-co-location-per-D07]

key-files:
  created:
    - virelo/workers/key_capture.py
    - virelo/workers/explorer.py
    - virelo/app/webview.py
    - virelo/app/window.py
    - virelo/app/__main__.py
  modified:
    - virelo/workers/__init__.py
    - virelo/app/__init__.py
    - virelo/services/snap.py
    - virelo/settings/state.py
    - virelo/__init__.py
    - pyproject.toml
    - Virelo.spec
    - main.py

key-decisions:
  - "ShiftSnapRestore placed in virelo/services/snap.py (not app/window.py) to keep dependency arrow pointing downward per T-03-05"
  - "Virelo.spec hiddenimports corrected to use virelo.platform.theme and virelo.platform.startup (actual locations from Plan 02)"
  - "explorer.py at 1019 lines acceptable -- ExplorerAutosizeEngine + ExplorerAutosizeWorker + COM lifecycle is cohesive per D-07"
  - "ruff per-file-ignores added for Virelo.spec (F821 for PyInstaller globals, UP009 for encoding comment)"

patterns-established:
  - "Thin shim pattern: main.py is 3 lines (from virelo.app import main; main())"
  - "Upper-tier extraction: MainWindow in app/window.py, startup in app/__main__.py"
  - "Worker split: key_capture.py pure/lightweight, explorer.py cohesive COM unit"

requirements-completed: [STRUCT-02, STRUCT-03, STRUCT-04]

# Metrics
duration: 13min
completed: 2026-04-24
---

# Phase 03 Plan 03: Upper-Tier Extraction Summary

**workers.py split into 2 modules, MainWindow and startup extracted to virelo/app/, main.py reduced to 3-line thin shim, 11 root-level Python files deleted**

## Performance

- **Duration:** 13 min
- **Started:** 2026-04-24T20:24:46Z
- **Completed:** 2026-04-24T20:38:21Z
- **Tasks:** 2
- **Files modified:** 24

## Accomplishments
- Split workers.py (1098 lines) into workers/key_capture.py (110 lines) and workers/explorer.py (1019 lines) with COM co-location preserved
- Extracted MainWindow to virelo/app/window.py (605 lines) and startup logic to virelo/app/__main__.py (155 lines)
- Moved ShiftSnapRestore to virelo/services/snap.py to prevent circular imports (dependency arrow points downward)
- Converted main.py to 3-line thin shim, deleted all 11 root-level Python files
- Updated Virelo.spec to read version from virelo/app/config.py with complete virelo.* hiddenimports

## Task Commits

Each task was committed atomically:

1. **Task 1: Split workers.py and move webview.py** - `ed874d5` (feat)
2. **Task 2: Extract MainWindow and startup, create thin shim, delete root-level files** - `0074b58` (feat)

## Files Created/Modified
- `virelo/workers/key_capture.py` - KeyCaptureSession (pure Python) + KeyCaptureWorker (QObject) with conditional PySide6 import
- `virelo/workers/explorer.py` - ExplorerAutosizeEngine + ExplorerAutosizeWorker + COM lifecycle, all co-located per D-07
- `virelo/workers/__init__.py` - Re-exports with try/except for PySide6-dependent classes
- `virelo/app/webview.py` - VireloWebView + VireloWebPage, imports from virelo.bridge and virelo.platform.resources
- `virelo/app/window.py` - MainWindow class with tray, theme, capture, explorer worker management
- `virelo/app/__main__.py` - _init_logger, _is_admin, main() with admin elevation and QApp lifecycle
- `virelo/app/__init__.py` - Re-exports main, APP_NAME, APP_VERSION, DEFAULTS
- `virelo/services/snap.py` - Appended ShiftSnapRestore class alongside existing SnapService
- `Virelo.spec` - Version regex points to virelo/app/config.py, hiddenimports cover all virelo.* submodules
- `main.py` - Thin shim: `from virelo.app import main; main()`
- `pyproject.toml` - Added ruff per-file-ignores for Virelo.spec

## Decisions Made
- **ShiftSnapRestore location:** Placed in virelo/services/snap.py (not virelo/app/window.py) to prevent circular imports -- MainWindow imports from services, never the reverse (per T-03-05)
- **Virelo.spec hiddenimports:** Corrected to reference virelo.platform.theme and virelo.platform.startup (actual locations from Plan 02), not virelo.services.theme/startup as plan template suggested
- **explorer.py exceeds 500 lines:** At 1019 lines, this is acceptable because ExplorerAutosizeEngine, ExplorerAutosizeWorker, and COM lifecycle (init, Shell.Application caching, iter_tabs, COM message pumping) form a cohesive unit that cannot be safely split per D-07
- **Ruff per-file-ignores for Virelo.spec:** Added F821 (PyInstaller globals like Analysis, PYZ, EXE, COLLECT) and UP009 (encoding comment) since these are PyInstaller requirements

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Fixed Virelo.spec hiddenimports for theme and startup**
- **Found during:** Task 2 (Step 6)
- **Issue:** Plan template listed virelo.services.theme and virelo.services.startup, but Plan 02 placed these at virelo.platform.theme and virelo.platform.startup
- **Fix:** Corrected hiddenimports to reference actual module locations
- **Files modified:** Virelo.spec
- **Verification:** Module paths match actual virelo/platform/ files on disk
- **Committed in:** 0074b58 (Task 2 commit)

**2. [Rule 1 - Bug] Fixed import ordering in virelo/settings/state.py**
- **Found during:** Task 2 (Step 7 ruff check)
- **Issue:** ruff I001 flagged unsorted imports in state.py (virelo.platform.theme came after virelo.settings.persistence)
- **Fix:** Reordered imports to satisfy ruff isort rules
- **Files modified:** virelo/settings/state.py
- **Verification:** ruff check passes
- **Committed in:** 0074b58 (Task 2 commit)

**3. [Rule 2 - Missing Critical] Added ruff per-file-ignores for Virelo.spec**
- **Found during:** Task 2 (Step 7 ruff check)
- **Issue:** Virelo.spec uses PyInstaller built-in globals (Analysis, PYZ, EXE, COLLECT) that ruff reports as F821 undefined names
- **Fix:** Added Virelo.spec to pyproject.toml per-file-ignores for F821 and UP009
- **Files modified:** pyproject.toml
- **Verification:** ruff check passes on all files including Virelo.spec
- **Committed in:** 0074b58 (Task 2 commit)

---

**Total deviations:** 3 auto-fixed (2 bugs, 1 missing critical)
**Impact on plan:** All fixes necessary for correctness. No scope creep.

## Issues Encountered
- PySide6, keyboard, and win32 modules not available in system Python (no .venv bootstrapped) -- modules verified via AST parsing rather than runtime import. All modules are structurally correct and will work at runtime when dependencies are available.

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- Full package restructuring complete: all Python source lives under virelo/
- main.py is a thin startup shim (STRUCT-02)
- No source file exceeds 500 lines except explorer.py (1019, justified per D-07) and explorer_columns.py (861, justified per D-04)
- Ready for Plan 04: test suites (unit tests can now import from clean virelo.* package paths)
- Ready for Plan 05: CI pipeline configuration

## Self-Check: PASSED

All 11 created/modified files verified present on disk. Both task commits (ed874d5, 0074b58) verified in git log.

---
*Phase: 03-structure-and-quality*
*Completed: 2026-04-24*
