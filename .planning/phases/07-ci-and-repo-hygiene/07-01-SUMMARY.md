---
phase: 07-ci-and-repo-hygiene
plan: 01
subsystem: build-tooling
tags: [git, powershell, clean-script, assets, repo-hygiene]

# Dependency graph
requires: []
provides:
  - icon.ico tracked by git (PyInstaller input)
  - branding/ tracked by git (Inno Setup installer graphics)
  - frontend/index.html tracked by git (Vite entry point)
  - scripts/clean.ps1 removes installer\dist and recursive __pycache__/\*.pyc
affects: [07-02, 07-03, 07-04]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "PowerShell Get-ChildItem -Recurse -Filter for recursive artifact removal"

key-files:
  created:
    - icon.ico
    - branding/installer-header.bmp
    - branding/installer-header_2x.bmp
    - branding/installer-wizard.bmp
    - branding/installer-wizard_2x.bmp
    - branding/virelo-icon.svg
    - frontend/index.html
  modified:
    - scripts/clean.ps1

key-decisions:
  - "Commit icon.ico, branding/*, and frontend/index.html as source assets (not generated artifacts) per D-01"
  - "Remove __pycache__ from flat targets list; replace with recursive Get-ChildItem pass per D-03"
  - "Add installer\\dist to flat targets list per D-04"

patterns-established:
  - "Pattern: Recursive pycache cleanup via Get-ChildItem -Recurse -Filter __pycache__ -Directory"

requirements-completed: [CI-01, CI-02]

# Metrics
duration: 1min
completed: 2026-04-25
---

# Phase 7 Plan 01: Source Asset Commit and Clean Script Expansion Summary

**Seven previously-untracked build inputs committed to git; clean.ps1 expanded to recursively remove all __pycache__ directories and *.pyc files plus installer\dist.**

## Performance

- **Duration:** 1min
- **Started:** 2026-04-25T01:43:01Z
- **Completed:** 2026-04-25T01:44:02Z
- **Tasks:** 2/2
- **Files modified:** 8 (7 new, 1 modified)

## Accomplishments

- Committed all 7 missing source assets required for PyInstaller and Inno Setup builds
- Expanded scripts/clean.ps1 to catch all __pycache__ directories anywhere in the project tree (not just root-level)
- Added installer\dist to the flat clean targets list

## Task Commits

Each task was committed atomically:

1. **Task 1: Stage missing source assets for git tracking** - `8efb6a9` (chore)
2. **Task 2: Expand scripts/clean.ps1 for recursive pycache and installer/dist** - `7f59bc7` (chore)

## Files Created/Modified

- `icon.ico` - Application icon required by PyInstaller spec (now tracked)
- `branding/installer-header.bmp` - Inno Setup installer header graphic (now tracked)
- `branding/installer-header_2x.bmp` - Inno Setup installer header graphic HiDPI (now tracked)
- `branding/installer-wizard.bmp` - Inno Setup installer wizard graphic (now tracked)
- `branding/installer-wizard_2x.bmp` - Inno Setup installer wizard graphic HiDPI (now tracked)
- `branding/virelo-icon.svg` - Source SVG icon (now tracked)
- `frontend/index.html` - Vite entry point required by frontend build (now tracked)
- `scripts/clean.ps1` - Expanded: removed __pycache__ from flat list, added installer\dist, added recursive Get-ChildItem passes

## Decisions Made

- Committed icon.ico, branding/*, and frontend/index.html as source inputs (not generated artifacts) — these are required for PyInstaller and Inno Setup, and were already present in the working tree
- Removed `__pycache__` from the `$targets` array; it only caught a root-level directory. Replaced with `Get-ChildItem -Recurse -Filter "__pycache__" -Directory` to catch all subdirectories in virelo/, virelo/app/, etc.
- Added `installer\dist` to the flat `$targets` array using backslash path separator consistent with existing `frontend\dist` entry

## Deviations from Plan

None - plan executed exactly as written.

## Self-Check: PASSED

- `git ls-files --error-unmatch icon.ico branding/installer-header.bmp branding/installer-header_2x.bmp branding/installer-wizard.bmp branding/installer-wizard_2x.bmp branding/virelo-icon.svg frontend/index.html` exits 0
- `installer\dist` found in scripts/clean.ps1 $targets array (line 10)
- `__pycache__` NOT in $targets array (lines 6-13)
- Two `Get-ChildItem -Recurse -Filter` passes present (lines 29 and 35)
- Commits `8efb6a9` and `7f59bc7` verified in git log

---
*Phase: 07-ci-and-repo-hygiene*
*Completed: 2026-04-25*
