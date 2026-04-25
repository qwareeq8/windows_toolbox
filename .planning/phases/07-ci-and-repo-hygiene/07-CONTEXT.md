# Phase 7: CI and Repo Hygiene - Context

**Gathered:** 2026-04-24
**Status:** Ready for planning

<domain>
## Phase Boundary

A developer can clone the repo and build, test, and lint successfully on the first try. This phase commits missing assets, fixes CI false positives and platform issues, synchronizes version strings, and renames stale class names.

</domain>

<decisions>
## Implementation Decisions

### Asset management (CI-01)
- **D-01:** Commit icon.ico, branding/*, and frontend/index.html to git. These are source assets required for building, not generated artifacts. They are currently untracked but present in the working tree.
- **D-02:** No .gitignore changes needed — frontend/dist/ is gitignored (build output), but frontend/index.html and branding/ are not blocked.

### Clean script expansion (CI-02)
- **D-03:** Expand scripts/clean.ps1 to recursively remove __pycache__/ and *.pyc across the entire project tree using Get-ChildItem -Recurse.
- **D-04:** Add installer/dist/ to the clean targets list.

### Stale-name false positive fix (CI-03)
- **D-05:** Add `--exclude-dir=.github` to the CI grep command so the workflow file doesn't match its own search string.
- **D-06:** Rephrase the test_app_config.py docstring (line 49) to avoid the literal "Windows Toolbox" string — use an indirect reference like "the old product name" instead.

### CI platform and test isolation (CI-04)
- **D-07:** Keep CI test runner on ubuntu-latest (free runners). Add `@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs")` to tests that import win32con/win32gui/ctypes (currently test_snap_geometry.py test_restore_maximized_window).
- **D-08:** Ensure all other unit tests pass on Ubuntu without platform-specific imports at module level. Use conditional imports guarded by sys.platform where needed.

### Version synchronization (CI-05, CI-06, CI-07)
- **D-09:** Sync frontend/package.json version to match config.py APP_VERSION (currently 1.4.2 vs 1.5.0 — update package.json to 1.5.0).
- **D-10:** Add a CI check step that extracts both versions and fails if they differ, preventing future drift.
- **D-11:** Update README.md "Python 3.8+" to "Python 3.12+" to match pyproject.toml requires-python = ">=3.12".
- **D-12:** Align installer support URL with config.py. Currently config.py uses `https://github.com/yusufqwareeq/virelo` while installer uses `mailto:qwareeq8@gmail.com`. Update installer MyAppURL to use the GitHub URL (more useful for support in Add/Remove Programs).

### Class renames (CI-08)
- **D-13:** Rename ShiftSnapRestore → SnapRestoreController in virelo/services/snap.py (class, docstrings, LOG messages), virelo/app/window.py (import, usage, comments), and tests/unit/test_snap_geometry.py (import, usage, comments).
- **D-14:** Rename HotkeyListener → MultiPressHotkeyListener in virelo/services/snap.py (class, docstrings) and virelo/app/window.py (import, usage).
- **D-15:** Update CLAUDE.md project structure section to reflect new class names.
- **D-16:** Do NOT rename references in .planning/ docs — those are historical records of the names at the time of writing.

### Claude's Discretion
- Exact pytest marker placement and import guard structure
- Clean script implementation details (PowerShell cmdlet choices)
- CI version-check script implementation (regex vs Python parse)
- Order of commits within the phase

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### CI workflow
- `.github/workflows/ci.yml` — Current CI pipeline: lint, test, frontend build, stale-name check
- `pyproject.toml` — Python version requirement, dev dependencies, Ruff config, pytest config

### Build pipeline
- `scripts/clean.ps1` — Current clean script (needs expansion)
- `scripts/build-frontend.ps1` — Frontend build script (version injection point)
- `scripts/build-app.ps1` — PyInstaller build script
- `scripts/build-installer.ps1` — Installer build script (passes version to ISCC)

### Version sources
- `virelo/app/config.py` — Single source of truth for APP_VERSION (1.5.0)
- `frontend/package.json` — Frontend version (currently 1.4.2 — out of sync)
- `installer/virelo.iss` — Installer version (received via /D flag at build time)

### Classes to rename
- `virelo/services/snap.py` — HotkeyListener (line 49) and ShiftSnapRestore (line 134) class definitions
- `virelo/app/window.py` — Import and usage of both classes
- `tests/unit/test_snap_geometry.py` — Test usage of ShiftSnapRestore (line 116)

### Project docs to update
- `CLAUDE.md` — Project structure references to snap.py classes
- `README.md` — Python version requirement (line says 3.8+, should be 3.12+)

No external specs — requirements are fully captured in decisions above.

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- scripts/clean.ps1: Existing clean script pattern (target list + loop) — extend, don't rewrite
- .github/workflows/ci.yml: 4-job CI structure (lint, test, frontend, stale-name) — add version-check as 5th job
- pyproject.toml: pytest config section available for adding markers/skip conditions

### Established Patterns
- PowerShell scripts use $ErrorActionPreference = "Stop" + $LASTEXITCODE checks
- CI uses ubuntu-latest with actions/setup-python@v5 and actions/setup-node@v4
- Version injection: build scripts read APP_VERSION from config.py via regex, pass to downstream tools
- Installer version uses `#ifndef` guard for /D override (CLAUDE.md footgun #4)

### Integration Points
- virelo/app/window.py line 23: `from virelo.services.snap import HotkeyListener, ShiftSnapRestore, SnapService` — rename point
- virelo/app/window.py lines 198-199: Instance creation of both classes — rename point
- SnapService.set_snap_mgr and set_listener methods accept renamed classes
- tests/unit/test_snap_geometry.py line 116: `from virelo.services.snap import ShiftSnapRestore` — rename point

</code_context>

<specifics>
## Specific Ideas

No specific requirements — open to standard approaches.

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope.

</deferred>

---

*Phase: 07-ci-and-repo-hygiene*
*Context gathered: 2026-04-24*
