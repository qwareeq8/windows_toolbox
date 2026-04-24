---
phase: 01-hygiene-and-build-pipeline
plan: 03
subsystem: infra
tags: [build-pipeline, powershell, pyinstaller, vite, inno-setup]

# Dependency graph
requires:
  - "01-01: app_config.py APP_VERSION, Virelo.spec, installer #ifndef guard, Vite define"
provides:
  - "Complete six-script PowerShell build pipeline from clean checkout to installer"
  - "bootstrap.ps1 for environment setup"
  - "clean.ps1 for build artifact cleanup"
  - "build-frontend.ps1 for React build with version injection"
  - "build-app.ps1 for PyInstaller build with frontend chaining"
  - "build-installer.ps1 rewritten with version passthrough to ISCC"
  - "verify-release.ps1 for post-build artifact validation"
affects: [01-hygiene-and-build-pipeline]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Build script chaining: build-installer -> build-app -> build-frontend"
    - "Version extraction via Select-String regex from app_config.py"
    - "VITE_APP_VERSION env var injection for frontend build"
    - "ISCC /DMyAppVersion flag for installer version"
    - "$LASTEXITCODE guard after every external command (PowerShell pitfall mitigation)"
    - "Precondition checks with Get-Command before using external tools"

key-files:
  created:
    - "scripts/bootstrap.ps1"
    - "scripts/clean.ps1"
    - "scripts/build-frontend.ps1"
    - "scripts/build-app.ps1"
    - "scripts/verify-release.ps1"
  modified:
    - "scripts/build-installer.ps1"

key-decisions:
  - "Build scripts chain via direct invocation (build-installer calls build-app calls build-frontend)"
  - "Each script validates preconditions and fails early with clear error messages"
  - "ISCC discovery uses ISCC_PATH env var with Program Files fallback candidates"

patterns-established:
  - "All build scripts use $ErrorActionPreference=Stop + $LASTEXITCODE guards"
  - "Version flows from app_config.py via Select-String regex in build scripts"
  - "Postcondition checks verify expected output files after each build step"

requirements-completed: [BUILD-01, BUILD-02, BUILD-03, BUILD-04, BUILD-05, BUILD-06]

# Metrics
duration: 2min
completed: 2026-04-24
---

# Phase 1 Plan 03: Build Pipeline Scripts Summary

**Six PowerShell build scripts implementing a full chained pipeline from clean checkout to versioned installer, with precondition validation and $LASTEXITCODE guards after every external command**

## Performance

- **Duration:** 2 min
- **Started:** 2026-04-24T18:08:11Z
- **Completed:** 2026-04-24T18:10:32Z
- **Tasks:** 3 (2 auto + 1 auto-approved checkpoint)
- **Files modified:** 6

## Accomplishments

- Created bootstrap.ps1: creates .venv, upgrades pip, installs requirements.txt deps, optionally installs npm deps
- Created clean.ps1: removes build/dist/frontend/dist/__pycache__/.pytest_cache/.ruff_cache and *.spec.bak files
- Created build-frontend.ps1: reads APP_VERSION from app_config.py, injects via VITE_APP_VERSION env var, runs npm build, verifies frontend/dist/index.html output
- Created build-app.ps1: validates .venv and Virelo.spec exist, chains build-frontend.ps1, runs PyInstaller, verifies dist/Virelo/Virelo.exe output
- Rewrote build-installer.ps1: chains build-app.ps1 (which chains build-frontend), reads APP_VERSION via Select-String, passes to ISCC via /DMyAppVersion flag, verifies installer/dist/VireloSetup.exe output
- Created verify-release.ps1: checks all expected artifacts exist without running any builds, detects stale Windows Toolbox.spec if present
- Every script starts with $ErrorActionPreference = "Stop" and guards all external commands with $LASTEXITCODE checks
- Every script that calls external tools (python, node, npm, pyinstaller, ISCC) validates their presence with Get-Command first

## Task Commits

Each task was committed atomically:

1. **Task 1: Create bootstrap.ps1, clean.ps1, build-frontend.ps1, and build-app.ps1** - `bc963a9` (feat)
2. **Task 2: Rewrite build-installer.ps1 and create verify-release.ps1** - `d9cb29f` (feat)
3. **Task 3: Checkpoint (human-verify)** - Auto-approved

## Files Created/Modified

- `scripts/bootstrap.ps1` - Virtual environment creation and dependency installation
- `scripts/clean.ps1` - Build artifact cleanup
- `scripts/build-frontend.ps1` - Frontend build with version injection and postcondition check
- `scripts/build-app.ps1` - PyInstaller build that calls build-frontend.ps1 first
- `scripts/build-installer.ps1` - Complete rewrite: chains build-app, injects version via ISCC /D flag
- `scripts/verify-release.ps1` - Post-build verification of all dist output artifacts

## Decisions Made

- Build scripts chain via direct invocation rather than a single monolithic script
- Each script validates preconditions and fails early with clear error messages
- ISCC discovery uses ISCC_PATH env var with Program Files fallback candidates

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness

- Complete build pipeline is in place: bootstrap -> build-frontend -> build-app -> build-installer
- Phase 1 is complete (all 3 plans executed): identity consolidation, repo hygiene, build pipeline
- Ready for Phase 2: bridge, security, and frontend work can build on this pipeline

## Self-Check: PASSED

All 6 script files verified present. Both task commits (bc963a9, d9cb29f) verified in git log.

---
*Phase: 01-hygiene-and-build-pipeline*
*Completed: 2026-04-24*
