---
phase: 08-window-chrome-and-release
verified: 2026-04-24T22:30:00Z
status: human_needed
score: 5/5 must-haves verified
overrides_applied: 0
human_verification:
  - test: "Drag the frameless window from the title bar area (top 35px excluding close/minimize buttons)"
    expected: "Window moves with the mouse without any dead zones in the title bar and without blocking close/minimize button clicks"
    why_human: "Visual/interactive behavior -- requires a running desktop environment to confirm WM_NCHITTEST hit zones produce correct drag behavior"
  - test: "Resize the window on a monitor with negative screen coordinates (e.g., secondary monitor to the left of primary)"
    expected: "Window resize grips respond correctly on all edges/corners without coordinate overflow"
    why_human: "Multi-monitor coordinate behavior cannot be verified without a multi-monitor physical setup"
  - test: "Run python -m virelo --smoke-test in activated venv"
    expected: "6 checks printed with PASS status, exit code 0, no UAC prompt shown"
    why_human: "Smoke test requires PySide6 and QWebEngine runtime which cannot be instantiated in a pure verification pass"
---

# Phase 8: Window Chrome and Release Verification Report

**Phase Goal:** The application window behaves correctly on all monitor configurations and the release pipeline produces a verified, documented artifact
**Verified:** 2026-04-24T22:30:00Z
**Status:** human_needed
**Re-verification:** No -- initial verification

## Goal Achievement

### Observable Truths

| #   | Truth | Status | Evidence |
| --- | ----- | ------ | -------- |
| 1   | User can drag the frameless window by clicking and dragging non-interactive title bar areas, using a Python-side event filter for reliable hit testing | VERIFIED | `virelo/app/window.py` line 498-501: returns HTCAPTION (2) for title bar zone (top 35px, excluding 4px border and 60px controls). TITLE_BAR_HEIGHT=35 matches frontend TitleBar height (34px + 1px border at app.jsx line 14-15). CONTROLS_WIDTH=60 excludes two 28px buttons + margin. Edge/corner checks take priority at lines 476-493. 22 unit tests pass covering all hit zones. |
| 2   | User can resize the window correctly even on monitors with negative screen coordinates (signed 16-bit lParam decoding) | VERIFIED | `virelo/app/window.py` lines 468-469: `ctypes.c_short(msg.lParam & 0xFFFF).value` for both x and y coordinates. Bare unsigned extraction (`msg.lParam & 0xFFFF`) no longer exists outside `c_short()`. 6 signed_short unit tests pass including -1920 round-trip and edge cases (-1, -32768, 32767). |
| 3   | Developer can run a non-interactive smoke test (--smoke-test) that verifies resource paths, frontend dist, QWebEngine init, settings, and bridge init without manual interaction | VERIFIED | `virelo/app/__main__.py` line 69: `_run_smoke_test()` function with 6 checks (icon.ico, frontend/dist/index.html, QWebEngine, Settings, SettingsState, VireloBridge). Lines 173-182: `--smoke-test` parsed via argparse before admin elevation (line 181 vs `_is_admin()` at line 208). Uses `parse_known_args` and `add_help=False` for Qt compatibility. Returns 0 on all pass, 1 on any fail. |
| 4   | Developer can run release verification that checks version consistency, asset existence, and built content correctness -- and .planning/ is gitignored while public docs exist in README.md and docs/ | VERIFIED | `scripts/verify-release.ps1` lines 52-82: 4 new checks (version cross-check via pkgJsonVersion, bundled icon.ico in dist/, bundled frontend/dist/index.html, stale naming scan in dist/). All use PowerShell cmdlets, follow existing `$errors` accumulation pattern. `.gitignore` line 50: `.planning/` entry present under "Planning artifacts (internal)" comment. `docs/` directory exists with BUILD.md, TROUBLESHOOTING.md, RELEASE.md. |
| 5   | Public documentation covers build instructions, troubleshooting, and release checklist with the correct product name and accurate descriptions | VERIFIED | `docs/BUILD.md`: covers prerequisites, quick start, full build pipeline table (bootstrap through verify-release), build order, dev mode, version management, smoke test. `docs/TROUBLESHOOTING.md`: covers 4 build issues and 4 runtime issues with causes and fixes. `docs/RELEASE.md`: 11-step release checklist from version update through git push. `README.md` lines 67-70: Documentation section links to all 3 docs files. No stale naming ("Windows Toolbox") found in docs/ or README.md (only correct reference in RELEASE.md describing the stale naming check itself). No hardcoded version strings in docs (uses x.y.z placeholders). |

**Score:** 5/5 truths verified

### Required Artifacts

| Artifact | Expected | Status | Details |
| -------- | -------- | ------ | ------- |
| `virelo/app/window.py` | HTCAPTION drag zone and signed lParam extraction in nativeEvent | VERIFIED | TITLE_BAR_HEIGHT=35, CONTROLS_WIDTH=60 constants at lines 32-33. Signed extraction at lines 468-469. HTCAPTION return at line 501. |
| `Virelo.spec` | UPX disabled for EXE and COLLECT | VERIFIED | `upx=False` at lines 67 and 81. No `upx=True` found anywhere. |
| `tests/unit/test_nchittest.py` | Unit tests for signed lParam extraction and hit-test zones | VERIFIED | 22 tests (6 signed_short + 16 classify_hit). All pass. |
| `virelo/app/__main__.py` | --smoke-test flag parsing and _run_smoke_test function | VERIFIED | argparse at line 173, _run_smoke_test at line 69, 6 checks, sys.exit before admin elevation. |
| `scripts/verify-release.ps1` | Expanded release verification with version cross-check, bundled asset checks, stale name scan | VERIFIED | Lines 52-82: pkgJsonVersion cross-check, dist/Virelo/icon.ico, dist/Virelo/frontend/dist/index.html, stale naming in dist/. |
| `.gitignore` | .planning/ exclusion entry | VERIFIED | Line 50: `.planning/` under "Planning artifacts (internal)" comment. |
| `docs/BUILD.md` | Build from source instructions for developers | VERIFIED | 82 lines covering prerequisites, quick start, build pipeline, dev mode, version management, smoke test. References `scripts/` correctly. |
| `docs/TROUBLESHOOTING.md` | Common issues and fixes | VERIFIED | 87 lines covering build issues (PyInstaller, PowerShell, Vite, Inno Setup) and runtime issues (UAC, drag, resize, explorer, smoke test). Uses "Virelo" exclusively. |
| `docs/RELEASE.md` | Release checklist | VERIFIED | 85 lines with 11-step checklist. References `verify-release.ps1`. Uses x.y.z version placeholders. |
| `README.md` | Updated entry point linking to docs/ | VERIFIED | Lines 67-70: Documentation section with links to docs/BUILD.md, docs/TROUBLESHOOTING.md, docs/RELEASE.md. |

### Key Link Verification

| From | To | Via | Status | Details |
| ---- | -- | --- | ------ | ------- |
| `virelo/app/window.py` | `frontend/src/app.jsx` TitleBar | TITLE_BAR_HEIGHT=35 matches TitleBar height(34)+border(1) | WIRED | window.py line 32: TITLE_BAR_HEIGHT=35. app.jsx line 14: height:34, line 15: borderBottom: 1px. 34+1=35. |
| `virelo/app/__main__.py` | `virelo/platform/resources.py` | resource_path('icon.ico') in smoke test check 1 | WIRED | __main__.py lines 97-99: imports and calls resource_path("icon.ico"). |
| `virelo/app/__main__.py` | `virelo/settings/persistence.py` | Settings() in smoke test check 4 | WIRED | __main__.py lines 126-130: imports Settings, instantiates, reads snap_key. |
| `virelo/app/__main__.py` | `virelo/settings/state.py` | SettingsState() in smoke test check 5 | WIRED | __main__.py lines 136-142: imports SettingsState, constructs with Settings, calls get_json. |
| `scripts/verify-release.ps1` | `virelo/app/config.py` | APP_VERSION regex extraction | WIRED | verify-release.ps1 line 11: Select-String extracting APP_VERSION. |
| `README.md` | `docs/BUILD.md` | link to detailed build docs | WIRED | README.md line 68: `[Building from Source](docs/BUILD.md)`. |
| `docs/BUILD.md` | `scripts/` | references to build scripts | WIRED | BUILD.md lines 28-33: full table of all scripts with descriptions. |
| `docs/RELEASE.md` | `scripts/verify-release.ps1` | references to release verification script | WIRED | RELEASE.md line 42: `scripts/verify-release.ps1`. |

### Data-Flow Trace (Level 4)

Not applicable for this phase. No artifacts render dynamic data from external sources. The phase modifies native event handling (Win32 messages from OS), CLI smoke testing (subsystem verification), build scripts (filesystem checks), and documentation (static text).

### Behavioral Spot-Checks

| Behavior | Command | Result | Status |
| -------- | ------- | ------ | ------ |
| Unit tests pass | `python -m pytest tests/unit/test_nchittest.py -x -v` | 22/22 passed | PASS |
| Full test suite regression | `python -m pytest tests/unit/ -q` | 88 passed in 0.05s | PASS |
| UPX disabled | Check `upx=True` not in Virelo.spec | No matches; `upx=False` appears exactly 2 times | PASS |
| Smoke test function exists | `grep "def _run_smoke_test" virelo/app/__main__.py` | Found at line 69 | PASS |
| Smoke test before admin | `grep -n "smoke_test\|_is_admin" virelo/app/__main__.py` | smoke_test at line 181, _is_admin at line 208 | PASS |
| .planning/ gitignored | `grep "^.planning/$" .gitignore` | Found at line 50 | PASS |
| All docs exist | `ls docs/` | BUILD.md, RELEASE.md, TROUBLESHOOTING.md | PASS |
| No stale naming in docs | `grep -ri "windows toolbox" docs/ README.md` | Only correct reference in RELEASE.md (describing what verify-release checks) | PASS |
| No hardcoded versions in docs | `grep "1.5.0" docs/*.md` | No matches | PASS |

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
| ----------- | ---------- | ----------- | ------ | -------- |
| CHRM-01 | 08-01 | User can drag frameless window from title bar areas via Python-side event filter | SATISFIED | HTCAPTION return in nativeEvent, TITLE_BAR_HEIGHT=35 matching frontend, CONTROLS_WIDTH=60 excluding buttons, edge/corner priority |
| CHRM-02 | 08-01 | User can resize on monitors with negative screen coordinates (signed 16-bit lParam) | SATISFIED | ctypes.c_short signed extraction for both x and y coordinates in nativeEvent |
| CHRM-03 | 08-01 | Developer builds without UPX compression | SATISFIED | upx=False on both EXE (line 67) and COLLECT (line 81) in Virelo.spec. Zero occurrences of upx=True. |
| CHRM-04 | 08-02 | Developer can run non-interactive smoke test (--smoke-test) | SATISFIED | _run_smoke_test function with 6 checks, parsed before admin elevation, exits 0/1 |
| CHRM-05 | 08-02 | Developer can run release verification for version consistency, assets, content | SATISFIED | 4 new checks in verify-release.ps1: version cross-check, bundled icon, bundled frontend, stale naming |
| CHRM-06 | 08-03 | .planning/ gitignored; public docs exist | SATISFIED | .planning/ in .gitignore line 50. docs/ directory with BUILD.md, TROUBLESHOOTING.md, RELEASE.md |
| CHRM-07 | 08-03 | Public docs cover build, troubleshooting, release with correct product name | SATISFIED | Three docs files with substantive content, all using "Virelo" exclusively, no hardcoded versions |

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
| ---- | ---- | ------- | -------- | ------ |
| (none) | - | - | - | No anti-patterns found in any modified files |

No TODO, FIXME, PLACEHOLDER, stub, or incomplete implementation patterns found in any phase-modified files (`virelo/app/window.py`, `virelo/app/__main__.py`, `Virelo.spec`, `scripts/verify-release.ps1`, `.gitignore`, `README.md`, `docs/BUILD.md`, `docs/TROUBLESHOOTING.md`, `docs/RELEASE.md`, `tests/unit/test_nchittest.py`).

### Human Verification Required

### 1. Title Bar Drag Behavior

**Test:** Launch Virelo, click and drag in the title bar area (top of window, left of minimize/close buttons)
**Expected:** Window follows the mouse smoothly. Close and minimize buttons remain clickable. Resize grips at all edges and corners still work. No dead zones in the drag area.
**Why human:** Visual/interactive behavior -- WM_NCHITTEST hit zones cannot be tested without a running desktop and mouse interaction.

### 2. Multi-Monitor Resize

**Test:** If a multi-monitor setup is available with a secondary monitor to the left or above the primary (negative coordinates), resize the Virelo window on that monitor from all edges and corners
**Expected:** Resize grips respond correctly, window follows the mouse, no coordinate wrapping or jumping
**Why human:** Requires a multi-monitor physical setup with negative screen coordinates to confirm signed lParam decoding works in practice.

### 3. Smoke Test Execution

**Test:** In an activated venv, run `python -m virelo --smoke-test`
**Expected:** Output shows "Virelo smoke test" header, 6 checks with "PASS" status, summary "6 passed, 0 failed", exit code 0, no UAC prompt
**Why human:** Smoke test requires PySide6 and QWebEngine runtime which cannot be verified in a static code analysis pass. Frontend dist must exist for check 2 to pass.

### Gaps Summary

No gaps found. All 5 roadmap success criteria are satisfied by substantive, wired code. All 7 requirements (CHRM-01 through CHRM-07) are covered with implementation evidence. All 22 unit tests pass. No regressions in the full test suite (88 tests). No anti-patterns detected. Three items require human verification for interactive/runtime behaviors that cannot be confirmed through static analysis.

---

_Verified: 2026-04-24T22:30:00Z_
_Verifier: Claude (gsd-verifier)_
