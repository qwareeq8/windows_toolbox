---
phase: 01-hygiene-and-build-pipeline
plan: 01
subsystem: infra
tags: [version, identity, pyinstaller, vite, inno-setup]

# Dependency graph
requires: []
provides:
  - "Single source of truth for product identity and version in app_config.py"
  - "Renamed Virelo.spec with regex-based version extraction"
  - "Inno Setup #ifndef guard for build-time version injection"
  - "Vite define block for __APP_VERSION__ frontend injection"
  - "Clean codebase with zero stale naming or migration references"
affects: [01-hygiene-and-build-pipeline]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Version flows from app_config.py via regex (spec), env var (Vite), /D flag (ISCC)"
    - "#ifndef guard pattern for Inno Setup build-time overrides"
    - "Vite define with JSON.stringify for build-time constant injection"

key-files:
  created:
    - "Virelo.spec"
  modified:
    - "app_config.py"
    - "installer/virelo.iss"
    - "frontend/vite.config.js"
    - "main.py"
    - "frontend/src/app.jsx"
    - "frontend/src/panels.jsx"
    - "frontend/src/pages.jsx"
    - "frontend/src/theme.jsx"
    - "frontend/src/icons.jsx"
    - "frontend/src/primitives.jsx"
    - "scripts/build-installer.ps1"

key-decisions:
  - "Version set to 1.5.0 as first release under consolidated identity"
  - "APP_LOG_DIR/APP_LOG_FILE added alongside existing LOG_DIR/LOG_FILE for backward compat"

patterns-established:
  - "Version single source of truth: app_config.py APP_VERSION"
  - "Spec reads version via regex, never imports app_config directly"
  - "Frontend displays version via __APP_VERSION__ Vite define global"

requirements-completed: [IDENT-01, IDENT-02, IDENT-03, IDENT-04, REPO-05, BUILD-04, BUILD-05]

# Metrics
duration: 5min
completed: 2026-04-24
---

# Phase 1 Plan 01: Product Identity and Version Consolidation Summary

**Consolidated all product identity into app_config.py, eliminated every "Windows Toolbox" reference, wired version flow through PyInstaller spec, Vite define, and Inno Setup #ifndef guard**

## Performance

- **Duration:** 5 min
- **Started:** 2026-04-24T17:54:33Z
- **Completed:** 2026-04-24T17:59:12Z
- **Tasks:** 2
- **Files modified:** 12

## Accomplishments
- APP_VERSION defined once in app_config.py as single source of truth, with 9 new product metadata constants
- Renamed Windows Toolbox.spec to Virelo.spec with regex-based version extraction (safe for PyInstaller context)
- Installer version uses #ifndef guard enabling build-time override via /D flag
- Vite config injects __APP_VERSION__ at build time via JSON.stringify
- All frontend version displays (app.jsx, panels.jsx, pages.jsx) now use __APP_VERSION__ global
- Zero stale "Windows Toolbox" or "Toolbox" references remain anywhere in source
- All migration-phase comments (Phase 7, SC-6, IC-11) cleaned from main.py
- Deprecated Qt 6 attributes (AA_EnableHighDpiScaling, AA_UseHighDpiPixmaps) removed
- All "v2" labels removed from frontend file comments

## Task Commits

Each task was committed atomically:

1. **Task 1: Extend app_config.py, rename spec, update installer and Vite config** - `524d851` (feat)
2. **Task 2: Remove stale naming, migration comments, deprecated Qt attrs, and wire frontend version** - `2bacf70` (fix)

## Files Created/Modified
- `app_config.py` - Added APP_VERSION, APP_DISPLAY_NAME, and 7 other D-02 metadata constants
- `Virelo.spec` - Renamed from Windows Toolbox.spec, added regex version extraction block
- `installer/virelo.iss` - Replaced hardcoded version with #ifndef MyAppVersion guard
- `frontend/vite.config.js` - Added define block for __APP_VERSION__ build-time injection
- `main.py` - Removed deprecated Qt attrs and migration-phase comments
- `frontend/src/app.jsx` - Replaced hardcoded version with __APP_VERSION__
- `frontend/src/panels.jsx` - Removed v2 label, replaced hardcoded version
- `frontend/src/pages.jsx` - Removed v2 label, replaced hardcoded version (preserved changelog data)
- `frontend/src/theme.jsx` - Removed v2 label from comment
- `frontend/src/icons.jsx` - Removed v2 label from comment
- `frontend/src/primitives.jsx` - Removed v2 label from comment
- `scripts/build-installer.ps1` - Updated spec reference from Windows Toolbox.spec to Virelo.spec

## Decisions Made
- Version set to 1.5.0 as the first release under consolidated identity (was 1.4.2)
- APP_LOG_DIR/APP_LOG_FILE added alongside existing LOG_DIR/LOG_FILE -- old names retained for backward compatibility with main.py imports

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Fixed build script referencing old spec filename**
- **Found during:** Task 2 (stale naming cleanup)
- **Issue:** `scripts/build-installer.ps1` referenced "Windows Toolbox.spec" in 3 places -- the plan's Task 2 file list omitted this file but the acceptance criteria required zero matches across .ps1 files
- **Fix:** Updated all 3 references to "Virelo.spec"
- **Files modified:** scripts/build-installer.ps1
- **Verification:** `grep -rn "Windows Toolbox" --include="*.ps1" .` returns zero matches
- **Committed in:** 2bacf70 (Task 2 commit)

---

**Total deviations:** 1 auto-fixed (1 bug)
**Impact on plan:** Essential fix -- build script would have failed without it. No scope creep.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- app_config.py is the single source of truth for version and product metadata
- Version flow is wired: app_config.py -> Virelo.spec (regex), -> Vite (env var), -> ISCC (/D flag)
- Build pipeline plans (01-02, 01-03) can now consume APP_VERSION from app_config.py
- Frontend version displays are dynamic -- will auto-update when VITE_APP_VERSION is set

---
*Phase: 01-hygiene-and-build-pipeline*
*Completed: 2026-04-24*
