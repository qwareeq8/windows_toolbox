# Phase 8: Window Chrome and Release - Research

**Researched:** 2026-04-24
**Domain:** Win32 window chrome (WM_NCHITTEST), PyInstaller packaging, smoke testing, release verification, documentation
**Confidence:** HIGH

## Summary

Phase 8 combines two distinct domains: window chrome correctness (frameless window dragging, signed coordinate decoding, UPX disabling) and release pipeline hardening (smoke test, expanded verification, public documentation). Both domains are well-defined with locked decisions from CONTEXT.md.

The window chrome work centers on extending the existing `nativeEvent` WM_NCHITTEST handler in `window.py` (lines 460-492). The existing handler already processes resize edges/corners but uses unsigned 16-bit lParam extraction, which breaks on monitors with negative screen coordinates. The fix is a two-line change using `ctypes.c_short`. The HTCAPTION addition for drag zones is a straightforward extension that returns hit code 2 for the top 35px of the window (matching the frontend TitleBar height of 34px plus 1px border), excluding resize borders and window control buttons.

The release pipeline work involves adding a `--smoke-test` flag to the application entry point, expanding the existing `verify-release.ps1` script, gitignoring `.planning/`, and creating a `docs/` directory with focused public documentation. All changes build on existing patterns and scripts.

**Primary recommendation:** Implement in three natural groupings: (1) window chrome fixes (CHRM-01, CHRM-02, CHRM-03), (2) smoke test and release verification (CHRM-04, CHRM-05), (3) documentation and gitignore (CHRM-06, CHRM-07).

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Extend the existing `nativeEvent` WM_NCHITTEST handler in `window.py` to return HTCAPTION (value 2) for mouse positions in the top title bar region. All hit-test logic stays in one method -- no separate event filter.
- **D-02:** The drag zone is the top ~40px of the window, excluding the 4px edge border zones (already handled for resize) and the right-side area where the frontend renders window control buttons (close, minimize). Python uses a fixed constant matching the frontend title bar height.
- **D-03:** Edge/corner resize zones (existing code) take priority -- the HTCAPTION check only fires if no edge/corner matched first.
- **D-04:** Replace the unsigned 16-bit extraction `x = msg.lParam & 0xFFFF` and `y = (msg.lParam >> 16) & 0xFFFF` with signed extraction using `ctypes.c_short` to handle monitors with negative screen coordinates. Use `x = ctypes.c_short(msg.lParam & 0xFFFF).value` and `y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value`.
- **D-05:** Set `upx=False` in both the EXE and COLLECT sections of `Virelo.spec`. UPX stays disabled until the release pipeline is verified stable. This is a two-line change.
- **D-06:** Add `--smoke-test` flag to `__main__.py` arg parsing. When present, run the full QApplication boot and MainWindow construction but do not call `app.exec()` or show the window. Verify each subsystem inline and report pass/fail.
- **D-07:** Smoke test checks: (1) icon.ico resource path resolves and exists, (2) frontend/dist/ directory exists and contains index.html, (3) QWebEngine can be constructed without errors, (4) Settings reads/writes without exceptions, (5) SettingsState initializes with valid defaults, (6) VireloBridge initializes without errors.
- **D-08:** Exit 0 if all checks pass, exit 1 if any fail. Print each check result to stdout for diagnostic output. No window is shown.
- **D-09:** Expand the existing `scripts/verify-release.ps1` rather than creating a new script. Add checks: (1) version in config.py matches frontend/package.json, (2) bundled icon.ico exists in dist/Virelo/, (3) bundled frontend/dist/ in dist/Virelo/ has index.html, (4) no stale "Windows Toolbox" naming in dist/ output.
- **D-10:** The script already checks dist/Virelo/Virelo.exe, installer output, and Virelo.spec. Keep those. Add the version cross-check and bundled content checks.
- **D-11:** Add `.planning/` to `.gitignore`. This directory contains internal planning artifacts not intended for public consumption.
- **D-12:** Create a `docs/` directory with public documentation. This replaces the planning artifacts for public-facing content.
- **D-13:** Keep README.md as the entry point with product overview, features, and quick start. Add a `docs/` directory with three focused files: `BUILD.md` (build instructions from source), `TROUBLESHOOTING.md` (common issues and fixes), and `RELEASE.md` (release checklist).
- **D-14:** All documentation uses "Virelo" product name. No references to "Windows Toolbox" or any stale naming.
- **D-15:** Build instructions cover the full pipeline: bootstrap, frontend build, app build, installer build, and release verification. Target audience is developers.

### Claude's Discretion
- Exact pixel height for the title bar drag zone constant (as long as it matches the frontend)
- How window control buttons are excluded from the drag zone (fixed pixel region vs. querying frontend)
- Smoke test output formatting (plain text, structured, or both)
- How to structure the README expansion (section ordering, level of detail)
- Whether docs/ files include screenshots or are text-only
- Order of commits within the phase

### Deferred Ideas (OUT OF SCOPE)
None -- analysis stayed within phase scope.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| CHRM-01 | User can drag the frameless window from non-interactive title bar areas via Python-side event filter | WM_NCHITTEST HTCAPTION pattern verified via Microsoft docs; TitleBar height measured at 34px + 1px border = 35px total; window control buttons are 28px wide each at right edge |
| CHRM-02 | User can resize the window correctly on monitors with negative screen coordinates (signed 16-bit lParam decoding) | Microsoft docs confirm lParam contains signed short values; ctypes.c_short conversion verified to correctly decode -1920 from unsigned 63616; existing code uses unsigned extraction which is the confirmed bug |
| CHRM-03 | Developer builds without UPX compression until release pipeline is verified stable | UPX+PySide6/Qt known to cause intermittent crashes; two-line change in Virelo.spec (lines 69, 81) |
| CHRM-04 | Developer can run a non-interactive smoke test (--smoke-test) | argparse integration in __main__.py; QApplication instantiation without app.exec(); subsystem verification pattern documented |
| CHRM-05 | Developer can run release verification that checks version consistency, asset existence, and built content correctness | Existing verify-release.ps1 has established patterns; version cross-check uses same regex as CI version-check job |
| CHRM-06 | .planning/ is gitignored and not published; public docs exist in README.md and docs/ | .gitignore format verified; docs/ directory confirmed not yet existing |
| CHRM-07 | Public documentation covers build instructions, troubleshooting, and release checklist with correct product name and accurate descriptions | Existing README.md structure reviewed; build pipeline documented in CLAUDE.md; all script names and behaviors verified |
</phase_requirements>

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Title bar drag (CHRM-01) | Python backend (Win32 native events) | Frontend (visual rendering) | WM_NCHITTEST is a Win32 message handled in Python; frontend only renders the visual title bar |
| Signed coordinate fix (CHRM-02) | Python backend (Win32 native events) | -- | Pure Python-side lParam decoding fix |
| UPX disabling (CHRM-03) | Build pipeline (Virelo.spec) | -- | PyInstaller configuration change |
| Smoke test (CHRM-04) | Python backend (entry point) | -- | Non-interactive application boot verification |
| Release verification (CHRM-05) | Build pipeline (PowerShell) | -- | Post-build artifact validation |
| Gitignore + docs (CHRM-06/07) | Repository config | -- | File management and documentation |

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| PySide6 | >=6.6 | Qt framework for nativeEvent WM_NCHITTEST handling | Already in use; nativeEvent is the standard approach for frameless window chrome [VERIFIED: pyproject.toml] |
| ctypes | stdlib | Signed 16-bit lParam extraction via c_short | Part of Python standard library; Microsoft docs recommend signed extraction [VERIFIED: Python docs, Microsoft docs] |
| argparse | stdlib | --smoke-test flag parsing | Standard library; used for simple flag parsing in entry points [VERIFIED: Python stdlib] |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| PyInstaller | >=6.0 | Build pipeline (UPX flag in spec) | Already in build deps; upx parameter documented [VERIFIED: pyproject.toml] |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| ctypes.c_short for lParam | struct.unpack('<hh', ...) | c_short is more readable for single value conversion; struct.unpack better for bulk parsing |
| argparse for --smoke-test | sys.argv manual check | argparse provides --help integration; for a single flag, sys.argv check is simpler but less extensible |

## Architecture Patterns

### System Architecture Diagram

```
User clicks title bar area
        |
        v
[Windows OS] ----WM_NCHITTEST (0x0084)----> [MainWindow.nativeEvent]
        |                                            |
        |                                    Extract x,y from lParam
        |                                    (signed via ctypes.c_short)
        |                                            |
        |                                    Priority checks:
        |                                    1. Resize edges/corners (BORDER=4px)
        |                                       -> return HTLEFT/HTTOP/etc.
        |                                    2. Title bar zone (top TITLE_BAR_HEIGHT px,
        |                                       excluding right CONTROLS_WIDTH px)
        |                                       -> return HTCAPTION (2)
        |                                    3. Default: fall through to super()
        |                                            |
        v                                            v
[Windows OS] interprets result:          [super().nativeEvent] handles
  HTCAPTION -> enables drag                remaining messages
  HTLEFT/etc -> enables resize
```

```
Developer runs: python -m virelo --smoke-test
        |
        v
[__main__.py] ----argparse----> --smoke-test flag detected
        |                               |
        |                       Skip admin elevation
        |                               |
        v                               v
[QApplication created]          [Smoke test runner]
        |                        Check 1: icon.ico exists
        |                        Check 2: frontend/dist/index.html
        |                        Check 3: QWebEngine constructs
        |                        Check 4: Settings reads/writes
        |                        Check 5: SettingsState defaults
        |                        Check 6: VireloBridge inits
        |                               |
        |                        Print results, exit 0 or 1
        v                        (no app.exec(), no window shown)
```

### Recommended Project Structure Changes
```
virelo/
  app/
    __main__.py          # Add --smoke-test flag, smoke_test() function
    window.py            # Extend nativeEvent with HTCAPTION, fix lParam
Virelo.spec              # upx=False on lines 69, 81
scripts/
  verify-release.ps1     # Expand with version cross-check, bundled asset checks
docs/                    # NEW directory
  BUILD.md               # Build from source instructions
  TROUBLESHOOTING.md     # Common issues and fixes
  RELEASE.md             # Release checklist
.gitignore               # Add .planning/ entry
README.md                # Expand with accurate product docs
```

### Pattern 1: WM_NCHITTEST with HTCAPTION for Drag Zones
**What:** Extend the existing nativeEvent handler to return HTCAPTION (2) for mouse positions in the title bar drag zone, after resize edge checks have had priority.
**When to use:** Frameless window needing title bar drag support without a native title bar.
**Example:**
```python
# Source: Microsoft WM_NCHITTEST docs + existing window.py pattern
# https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest

TITLE_BAR_HEIGHT = 35  # 34px TitleBar + 1px borderBottom
CONTROLS_WIDTH = 60    # Two 28px buttons + padding margin

def nativeEvent(self, event_type, message):
    if event_type == b"windows_generic_MSG":
        msg = ctypes.wintypes.MSG.from_address(int(message))
        if msg.message == 0x0084:  # WM_NCHITTEST
            # Signed extraction (handles negative monitor coords)
            x = ctypes.c_short(msg.lParam & 0xFFFF).value
            y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
            pos = self.mapFromGlobal(QtCore.QPoint(x, y))
            rect = self.rect()
            BORDER = 4

            # 1. Edges and corners (existing code) - PRIORITY
            result = 0
            if pos.x() <= BORDER:
                # ... existing edge logic ...
            # ... etc ...

            if result:
                return True, result

            # 2. Title bar drag zone (NEW)
            if (pos.y() < TITLE_BAR_HEIGHT
                    and pos.x() >= BORDER
                    and pos.x() < rect.width() - CONTROLS_WIDTH):
                return True, 2  # HTCAPTION

    return super().nativeEvent(event_type, message)
```
[VERIFIED: Microsoft WM_NCHITTEST docs, existing window.py code]

### Pattern 2: Signed lParam Extraction
**What:** Convert unsigned 16-bit values from lParam to signed using ctypes.c_short.
**When to use:** Any WM_NCHITTEST or WM_MOUSEMOVE handler that processes screen coordinates on multi-monitor systems.
**Example:**
```python
# Source: https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest
# "Do not use LOWORD or HIWORD macros... these macros return incorrect results
#  on systems with multiple monitors"

# BUGGY (unsigned -- breaks on negative coordinates):
x = msg.lParam & 0xFFFF           # Returns 63616 for x=-1920
y = (msg.lParam >> 16) & 0xFFFF

# CORRECT (signed -- works on all monitor configurations):
x = ctypes.c_short(msg.lParam & 0xFFFF).value           # Returns -1920
y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
```
[VERIFIED: Microsoft docs + local ctypes.c_short test confirmed -1920 round-trips correctly]

### Pattern 3: Non-Interactive Smoke Test
**What:** Add a --smoke-test argument to the application entry point that boots the Qt application, verifies subsystems, and exits without entering the event loop.
**When to use:** CI/CD or local pre-release verification of application initialization without manual interaction.
**Example:**
```python
# Source: PySide6 QApplication docs + argparse stdlib docs

import argparse

def main():
    parser = argparse.ArgumentParser(prog="virelo")
    parser.add_argument("--smoke-test", action="store_true",
                        help="Run non-interactive boot verification and exit")
    args, _ = parser.parse_known_args()

    if args.smoke_test:
        return _run_smoke_test()

    # ... normal application startup ...

def _run_smoke_test():
    """Boot QApplication and verify subsystems without entering event loop."""
    import sys
    from PySide6 import QtWidgets
    app = QtWidgets.QApplication(sys.argv)
    passed = 0
    failed = 0

    def check(name, fn):
        nonlocal passed, failed
        try:
            fn()
            print(f"  PASS: {name}")
            passed += 1
        except Exception as e:
            print(f"  FAIL: {name} -- {e}")
            failed += 1

    print("Virelo smoke test")
    print("=" * 40)
    check("icon.ico exists", lambda: ...)
    # ... more checks ...
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(0 if failed == 0 else 1)
```
[ASSUMED: smoke test pattern based on standard PySide6 usage; no QWebEngine-specific smoke test documentation found]

### Anti-Patterns to Avoid
- **Calling `app.exec()` in smoke test:** The event loop will block forever. The smoke test must verify subsystem init and exit before entering the loop.
- **Unsigned lParam extraction on multi-monitor:** The existing code (`msg.lParam & 0xFFFF`) treats coordinates as unsigned, producing garbage values (63616 instead of -1920) when monitors have negative screen coordinates.
- **Separate event filter for drag:** D-01 explicitly requires all hit-test logic in one method. Do not create a separate `installEventFilter` handler.
- **Hardcoding title bar height without matching frontend:** The Python constant must match the frontend TitleBar's actual height (34px content + 1px border = 35px). If the frontend changes, the constant must be updated.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Signed 16-bit extraction | Manual bit-shifting with sign extension | `ctypes.c_short(val).value` | c_short handles two's complement correctly for all edge cases |
| CLI argument parsing | Manual sys.argv parsing | `argparse` with `parse_known_args` | Handles --help, future flags, and ignores Qt arguments |
| Version string parsing in PowerShell | Custom regex patterns | `Select-String -Pattern` (existing pattern) | Already used in verify-release.ps1; consistent with codebase |

**Key insight:** All the components being modified (nativeEvent handler, verify-release.ps1, __main__.py) already exist with established patterns. This phase extends existing code rather than creating new systems.

## Common Pitfalls

### Pitfall 1: HTCAPTION Swallowing Button Clicks
**What goes wrong:** Returning HTCAPTION for the entire top region prevents the React frontend's minimize/close buttons from receiving click events, because Windows treats HTCAPTION clicks as drag operations.
**Why it happens:** The WM_NCHITTEST response occurs before the mouse event reaches the web content. If Python returns HTCAPTION over the button area, the click never arrives at the React button.
**How to avoid:** Exclude the rightmost area (where minimize and close buttons live) from the HTCAPTION zone. The two buttons are each 28px wide (total ~60px from right edge). Any position with `pos.x() >= rect.width() - CONTROLS_WIDTH` must NOT return HTCAPTION.
**Warning signs:** Minimize and close buttons stop responding to clicks after the drag zone is added.

### Pitfall 2: Title Bar Height Mismatch
**What goes wrong:** The Python TITLE_BAR_HEIGHT constant does not match the frontend's actual rendered title bar height, causing either a dead zone (too small) or drag interfering with content below the title bar (too large).
**Why it happens:** The frontend TitleBar height is set in JavaScript (app.jsx line 14: `height: 34`) plus a 1px borderBottom. If either side changes independently, they desync.
**How to avoid:** Define the constant clearly with a comment referencing the frontend source. The current frontend TitleBar is 34px + 1px border = 35px total. The CONTEXT.md says "~40px" which gives 5px of margin; using the exact 35px measured from frontend source is more precise.
**Warning signs:** Small gap between title bar bottom and drag-responsive area, or drag starts interfering with the search bar below the title.

### Pitfall 3: Smoke Test Admin Elevation Loop
**What goes wrong:** The smoke test triggers the UAC elevation prompt, which re-launches the process, which tries to elevate again, creating an infinite loop or requiring admin rights for a simple test.
**Why it happens:** The existing `__main__.py` calls `_is_admin()` early and re-launches via `ShellExecuteW("runas", ...)` if not admin. If `--smoke-test` is parsed after the elevation check, the smoke test always requires admin.
**How to avoid:** Parse `--smoke-test` BEFORE the admin elevation check. If `--smoke-test` is present, skip elevation entirely and proceed directly to the smoke test function. The smoke test verifies initialization, not runtime behavior that requires admin privileges.
**Warning signs:** UAC prompt appears when running `python main.py --smoke-test`.

### Pitfall 4: PowerShell $LASTEXITCODE Not Checked
**What goes wrong:** New checks added to verify-release.ps1 use external commands (like `Select-String`) that might fail silently.
**Why it happens:** CLAUDE.md footgun #2: `$ErrorActionPreference = "Stop"` only catches PowerShell cmdlet errors, not native command failures.
**How to avoid:** The existing verify-release.ps1 already uses `$errors` accumulation pattern (not external commands). New checks should follow the same pattern: use PowerShell cmdlets (`Test-Path`, `Select-String`, `Get-Content`) and accumulate failures into `$errors`. No external commands needed for the new checks.
**Warning signs:** verify-release.ps1 reports PASSED when it should have caught a failure.

### Pitfall 5: UPX Change Breaks Spec Parsing
**What goes wrong:** Editing Virelo.spec to change `upx=True` to `upx=False` accidentally introduces syntax errors or changes the wrong line.
**Why it happens:** The spec file has `upx=True` on two separate lines (EXE at line 69, COLLECT at line 81). Missing one leaves UPX partially enabled.
**How to avoid:** Both occurrences must be changed. The lines are in different blocks (EXE and COLLECT). Verify both by searching for `upx=` after the edit.
**Warning signs:** Build succeeds but some binaries are still compressed.

## Code Examples

### WM_NCHITTEST Extension (Complete)
```python
# Source: Verified from Microsoft docs + existing window.py (lines 460-492)
# https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest

# Constants at class or module level
TITLE_BAR_HEIGHT = 35   # Frontend TitleBar: 34px height + 1px borderBottom
CONTROLS_WIDTH = 60     # Two window control buttons: 28px each + safety margin

def nativeEvent(self, event_type, message):
    if event_type == b"windows_generic_MSG":
        msg = ctypes.wintypes.MSG.from_address(int(message))
        if msg.message == 0x0084:  # WM_NCHITTEST
            # Signed extraction for multi-monitor support (D-04)
            x = ctypes.c_short(msg.lParam & 0xFFFF).value
            y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value
            pos = self.mapFromGlobal(QtCore.QPoint(x, y))
            rect = self.rect()
            BORDER = 4  # 4px grab zone

            result = 0
            # Edges and corners (existing, unchanged)
            if pos.x() <= BORDER:
                if pos.y() <= BORDER:
                    result = 13  # HTTOPLEFT
                elif pos.y() >= rect.height() - BORDER:
                    result = 16  # HTBOTTOMLEFT
                else:
                    result = 10  # HTLEFT
            elif pos.x() >= rect.width() - BORDER:
                if pos.y() <= BORDER:
                    result = 14  # HTTOPRIGHT
                elif pos.y() >= rect.height() - BORDER:
                    result = 17  # HTBOTTOMRIGHT
                else:
                    result = 11  # HTRIGHT
            elif pos.y() <= BORDER:
                result = 12  # HTTOP
            elif pos.y() >= rect.height() - BORDER:
                result = 15  # HTBOTTOM

            if result:
                return True, result

            # Title bar drag zone (D-01, D-02, D-03)
            if (pos.y() < TITLE_BAR_HEIGHT
                    and pos.x() < rect.width() - CONTROLS_WIDTH):
                return True, 2  # HTCAPTION

    return super().nativeEvent(event_type, message)
```
[VERIFIED: Microsoft WM_NCHITTEST docs confirm all hit codes; existing window.py code verified]

### Smoke Test Structure
```python
# Source: argparse stdlib + PySide6 patterns
# Integration point: virelo/app/__main__.py

def _run_smoke_test():
    """Non-interactive boot verification (D-06, D-07, D-08)."""
    import sys
    from PySide6 import QtCore, QtWidgets
    from virelo.app.config import APP_NAME, ORGANIZATION

    QtCore.QCoreApplication.setOrganizationName(ORGANIZATION)
    QtCore.QCoreApplication.setApplicationName(APP_NAME)
    app = QtWidgets.QApplication(sys.argv)

    passed = 0
    failed = 0

    def check(name, fn):
        nonlocal passed, failed
        try:
            result = fn()
            if result is False:
                raise AssertionError("returned False")
            print(f"  PASS  {name}")
            passed += 1
        except Exception as e:
            print(f"  FAIL  {name} -- {e}")
            failed += 1

    print(f"Virelo smoke test")
    print("=" * 40)

    # Check 1: icon.ico
    check("icon.ico resource path", lambda: (
        __import__('os').path.exists(
            __import__('virelo.platform.resources', fromlist=['resource_path']).resource_path("icon.ico")
        ) or (_ for _ in ()).throw(FileNotFoundError("icon.ico not found"))
    ))

    # Check 2-6: more checks following same pattern...

    print(f"\n{passed} passed, {failed} failed")
    sys.exit(0 if failed == 0 else 1)
```
[ASSUMED: smoke test structure; actual implementation should use clearer lambda bodies]

### Verify-Release Expansion
```powershell
# Source: existing scripts/verify-release.ps1 pattern

# --- Version cross-check (D-09 item 1) ---
$pkgJsonVersion = (Get-Content "frontend\package.json" | ConvertFrom-Json).version
if ($AppVersion -ne $pkgJsonVersion) {
    $errors += "Version mismatch: config.py=$AppVersion, package.json=$pkgJsonVersion"
} else {
    Write-Host "[verify-release] OK: Versions match ($AppVersion)"
}

# --- Bundled icon.ico in dist/ (D-09 item 2) ---
if (-not (Test-Path "dist\Virelo\icon.ico")) {
    $errors += "Missing: dist\Virelo\icon.ico"
} else {
    Write-Host "[verify-release] OK: dist/Virelo/icon.ico"
}

# --- Bundled frontend in dist/ (D-09 item 3) ---
if (-not (Test-Path "dist\Virelo\frontend\dist\index.html")) {
    $errors += "Missing: dist\Virelo\frontend\dist\index.html"
} else {
    Write-Host "[verify-release] OK: dist/Virelo/frontend/dist/index.html"
}

# --- Stale naming in dist/ (D-09 item 4) ---
$staleMatch = Get-ChildItem "dist\Virelo" -Recurse -File |
    Where-Object { $_.Name -match "Windows.Toolbox" -or $_.Name -match "toolbox" }
if ($staleMatch) {
    $errors += "Stale naming found in dist/: $($staleMatch.Name -join ', ')"
}
```
[VERIFIED: existing verify-release.ps1 pattern; PowerShell cmdlets used per CLAUDE.md footgun #2]

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| LOWORD/HIWORD macros for lParam | GET_X_LPARAM/GET_Y_LPARAM (or ctypes.c_short in Python) | Always -- Microsoft docs have always warned about this | Prevents negative coordinate bugs on multi-monitor |
| UPX compression for Qt apps | Skip UPX for Qt/PySide6 binaries | PyInstaller disables UPX on non-Windows by default since 2023 | Prevents intermittent startup crashes |

**Deprecated/outdated:**
- UPX compression for PySide6 apps: Known to cause intermittent crashes due to Qt DLL decompression timing issues [CITED: PyInstaller GitHub issue #4178, discussion #8922]

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | Frontend TitleBar height is 35px total (34px + 1px border) | Architecture Patterns, Pitfall 2 | Drag zone will be off by a few pixels -- functional but imprecise |
| A2 | Window control buttons occupy ~60px from right edge (2x 28px + margin) | Architecture Patterns, Pitfall 1 | Buttons could be blocked by HTCAPTION or have dead zone between drag area and buttons |
| A3 | Smoke test can construct QWebEngine without entering the event loop | Pattern 3 | QWebEngine may require event loop processing to fully initialize; test may need to process events briefly |
| A4 | --smoke-test should skip admin elevation | Pitfall 3 | If smoke test checks require admin access, skipping elevation would cause false failures |

## Open Questions (RESOLVED)

1. **Exact CONTROLS_WIDTH value** — RESOLVED: Using 60px per Claude's discretion, adopted in Plan 08-01 Task 2.
   - What we know: Each button is 28px wide (from app.jsx lines 36-48). Two buttons = 56px.
   - What's unclear: Whether there is padding between the last button and the window edge, or between buttons. The buttons use no explicit gap in the flex container, but the TitleBar has `padding: '0 6px 0 12px'` which adds 6px on the right.
   - Resolution: Use 60px as CONTROLS_WIDTH (56px buttons + 4px safety margin). This is generous enough to avoid the pitfall without creating a noticeable dead zone.

2. **Smoke test and QWebEngine initialization** — RESOLVED: Construct without event loop, adopted in Plan 08-02 Task 1.
   - What we know: D-07 specifies "QWebEngine can be constructed without errors" as a check.
   - What's unclear: Whether constructing a QWebEngineView without entering the event loop will succeed or throw. VireloWebView constructor (webview.py) does `self.setUrl(url)` which may require event processing.
   - Resolution: Construct the QWebEngineView but do not verify URL loading. Checking that the constructor does not throw is sufficient for a smoke test.

3. **Smoke test and admin rights** — RESOLVED: Skip elevation for smoke test, adopted in Plan 08-02 Task 1.
   - What we know: The app normally requires admin for keyboard hooks and cross-process window manipulation.
   - What's unclear: Whether Settings (QSettings), SettingsState, or VireloBridge initialization require admin rights.
   - Resolution: QSettings, SettingsState, and VireloBridge are pure in-process operations that do not require admin. Skip elevation for smoke test.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Python | All | Yes | 3.13.13 | -- |
| Node.js | Frontend build check | Yes | 24.15.0 | -- |
| npm | Frontend build check | Yes | 11.12.1 | -- |
| PySide6 | Smoke test, window chrome | In .venv (not system Python) | >=6.6 (per pyproject.toml) | Must run from activated venv |
| PyInstaller | UPX flag change | In dev deps | >=6.0 (per pyproject.toml) | -- |

**Missing dependencies with no fallback:** None.

**Missing dependencies with fallback:** None.

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | pytest >=9.0.3 (Python), vitest >=4.1.5 (Frontend) |
| Config file | pyproject.toml [tool.pytest.ini_options] |
| Quick run command | `pytest tests/unit/ -q` |
| Full suite command | `pytest tests/unit/ -q && cd frontend && npx vitest run` |

### Phase Requirements to Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| CHRM-01 | HTCAPTION returned for title bar positions | unit | `pytest tests/unit/test_nchittest.py -x` | No -- Wave 0 |
| CHRM-02 | Signed lParam decoding for negative coords | unit | `pytest tests/unit/test_nchittest.py -x` | No -- Wave 0 |
| CHRM-03 | upx=False in Virelo.spec | smoke | `python -c "import re; s=open('Virelo.spec').read(); assert 'upx=True' not in s"` | No inline check |
| CHRM-04 | --smoke-test exits 0 on healthy env | integration | `python -m virelo --smoke-test` (requires PySide6/venv) | No -- new feature |
| CHRM-05 | verify-release.ps1 catches mismatches | manual-only | `scripts/verify-release.ps1` (requires built artifacts) | Yes (script exists) |
| CHRM-06 | .planning/ in .gitignore | smoke | `grep -q '.planning/' .gitignore` | No inline check |
| CHRM-07 | docs/ files exist with correct content | smoke | `test -f docs/BUILD.md && test -f docs/TROUBLESHOOTING.md && test -f docs/RELEASE.md` | No -- new files |

### Sampling Rate
- **Per task commit:** `pytest tests/unit/ -q`
- **Per wave merge:** `pytest tests/unit/ -q && cd frontend && npx vitest run`
- **Phase gate:** Full suite green before verify

### Wave 0 Gaps
- [ ] `tests/unit/test_nchittest.py` -- unit tests for signed lParam extraction and HTCAPTION hit zones (covers CHRM-01, CHRM-02)
- [ ] No test framework gaps -- pytest already configured and functional

## Security Domain

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | No | N/A -- desktop utility, no auth |
| V3 Session Management | No | N/A -- no sessions |
| V4 Access Control | No | N/A -- single user |
| V5 Input Validation | Yes | Smoke test flag validated by argparse; WM_NCHITTEST coordinates validated by bounds checks |
| V6 Cryptography | No | N/A -- no crypto operations |

### Known Threat Patterns for Win32 Chrome

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| Malicious WM_NCHITTEST coordinates | Tampering | Coordinates are from OS; validated by mapFromGlobal + bounds checking |
| Smoke test as attack vector | Information Disclosure | --smoke-test only prints PASS/FAIL for subsystems; no sensitive data exposed |

## Project Constraints (from CLAUDE.md)

- **Forbidden:** Never reintroduce "Windows Toolbox" or "Toolbox" in any file -- applies to all documentation files (README.md, docs/*.md)
- **Forbidden:** Never add fake or placeholder UI controls -- no changes to UI in this phase, but smoke test must not create fake controls
- **Forbidden:** Never commit generated artifacts -- docs/ files are hand-written, not generated
- **Forbidden:** Never hardcode version strings -- README.md and docs/ must reference version via config.py or the build pipeline, not literal version numbers
- **Footgun #1:** config.py must not be imported in Virelo.spec -- UPX change is a literal text edit, no import involved
- **Footgun #2:** PowerShell $LASTEXITCODE -- verify-release.ps1 uses cmdlets, not external commands, for new checks
- **Footgun #3:** Vite define values -- not relevant to this phase
- **Footgun #4:** Inno Setup #define guard -- not relevant to this phase

## Sources

### Primary (HIGH confidence)
- [Microsoft WM_NCHITTEST documentation](https://learn.microsoft.com/en-us/windows/win32/inputdev/wm-nchittest) -- lParam format, signed coordinates, HTCAPTION value (2), all hit-test return codes
- `virelo/app/window.py` (lines 460-492) -- existing nativeEvent handler with unsigned lParam extraction
- `frontend/src/app.jsx` (lines 10-51) -- TitleBar component: height 34px, minimize/close buttons 28px each
- `virelo/app/__main__.py` (lines 68-155) -- existing entry point with admin elevation
- `Virelo.spec` (lines 69, 81) -- upx=True in EXE and COLLECT
- `scripts/verify-release.ps1` -- existing verification script pattern

### Secondary (MEDIUM confidence)
- [PyInstaller GitHub Discussion #8922](https://github.com/orgs/pyinstaller/discussions/8922) -- UPX disabled on non-Windows by default
- [PyInstaller GitHub Issue #4178](https://github.com/pyinstaller/pyinstaller/issues/4178) -- UPX crashes at startup with Qt binaries
- Local ctypes.c_short verification -- confirmed -1920 round-trips through unsigned 16-bit encoding

### Tertiary (LOW confidence)
- None -- all claims verified against primary or secondary sources

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- all libraries already in use, no new dependencies
- Architecture: HIGH -- extending existing patterns with well-documented Win32 APIs
- Pitfalls: HIGH -- pitfalls derived from direct code inspection and Microsoft documentation
- Smoke test: MEDIUM -- QWebEngine behavior without event loop is assumed, not verified

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (stable -- Win32 API and existing codebase patterns)
