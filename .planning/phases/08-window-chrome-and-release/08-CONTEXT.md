# Phase 8: Window Chrome and Release - Context

**Gathered:** 2026-04-24 (auto mode)
**Status:** Ready for planning

<domain>
## Phase Boundary

The application window behaves correctly on all monitor configurations and the release pipeline produces a verified, documented artifact. This phase adds frameless window dragging, fixes signed coordinate decoding, disables UPX until verified, adds a non-interactive smoke test, expands release verification, gitignores .planning/, and creates public documentation.

</domain>

<decisions>
## Implementation Decisions

### Title bar drag zones (CHRM-01)
- **D-01:** Extend the existing `nativeEvent` WM_NCHITTEST handler in `window.py` to return HTCAPTION (value 2) for mouse positions in the top title bar region. All hit-test logic stays in one method — no separate event filter.
- **D-02:** The drag zone is the top ~40px of the window, excluding the 4px edge border zones (already handled for resize) and the right-side area where the frontend renders window control buttons (close, minimize). Python uses a fixed constant matching the frontend title bar height.
- **D-03:** Edge/corner resize zones (existing code) take priority — the HTCAPTION check only fires if no edge/corner matched first.

### Signed coordinate decoding (CHRM-02)
- **D-04:** Replace the unsigned 16-bit extraction `x = msg.lParam & 0xFFFF` and `y = (msg.lParam >> 16) & 0xFFFF` with signed extraction using `ctypes.c_short` to handle monitors with negative screen coordinates. Use `x = ctypes.c_short(msg.lParam & 0xFFFF).value` and `y = ctypes.c_short((msg.lParam >> 16) & 0xFFFF).value`.

### UPX compression (CHRM-03)
- **D-05:** Set `upx=False` in both the EXE and COLLECT sections of `Virelo.spec`. UPX stays disabled until the release pipeline is verified stable. This is a two-line change.

### Smoke test (CHRM-04)
- **D-06:** Add `--smoke-test` flag to `__main__.py` arg parsing. When present, run the full QApplication boot and MainWindow construction but do not call `app.exec()` or show the window. Verify each subsystem inline and report pass/fail.
- **D-07:** Smoke test checks: (1) icon.ico resource path resolves and exists, (2) frontend/dist/ directory exists and contains index.html, (3) QWebEngine can be constructed without errors, (4) Settings reads/writes without exceptions, (5) SettingsState initializes with valid defaults, (6) VireloBridge initializes without errors.
- **D-08:** Exit 0 if all checks pass, exit 1 if any fail. Print each check result to stdout for diagnostic output. No window is shown.

### Release verification (CHRM-05)
- **D-09:** Expand the existing `scripts/verify-release.ps1` rather than creating a new script. Add checks: (1) version in config.py matches frontend/package.json, (2) bundled icon.ico exists in dist/Virelo/, (3) bundled frontend/dist/ in dist/Virelo/ has index.html, (4) no stale "Windows Toolbox" naming in dist/ output.
- **D-10:** The script already checks dist/Virelo/Virelo.exe, installer output, and Virelo.spec. Keep those. Add the version cross-check and bundled content checks.

### .planning/ gitignore (CHRM-06)
- **D-11:** Add `.planning/` to `.gitignore`. This directory contains internal planning artifacts not intended for public consumption.
- **D-12:** Create a `docs/` directory with public documentation. This replaces the planning artifacts for public-facing content.

### Public documentation (CHRM-07)
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

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Vision, constraints, out-of-scope features, key decisions
- `.planning/REQUIREMENTS.md` — CHRM-01..07 requirement definitions (Phase 8 scope)
- `CLAUDE.md` — Forbidden changes (no fake controls, no stale naming, no committed artifacts), known footguns (spec import, $LASTEXITCODE, Vite define, Inno #ifndef)

### Window chrome
- `virelo/app/window.py` — MainWindow: nativeEvent WM_NCHITTEST handler (lines 460-492), FramelessWindowHint (line 147), center_on_screen (line 445)
- `virelo/platform/win32_helpers.py` — Win32 utility functions, DPI awareness, monitor rect calculations

### Build and release
- `Virelo.spec` — PyInstaller spec: upx=True (lines 69, 81), datas, hiddenimports
- `scripts/verify-release.ps1` — Existing release verification (version, artifacts, stale name checks)
- `scripts/build-app.ps1` — PyInstaller build script
- `scripts/build-installer.ps1` — Installer build script
- `virelo/app/config.py` — APP_VERSION single source of truth

### Entry point
- `virelo/app/__main__.py` — Application entry point: logging, admin elevation, single instance, QApp creation, MainWindow (--smoke-test addition point)

### Prior phase context
- `.planning/phases/07-ci-and-repo-hygiene/07-CONTEXT.md` — CI pipeline patterns, version sync (D-09..D-12), clean script patterns (D-03..D-04)
- `.planning/phases/05-bridge-and-settings-correctness/05-CONTEXT.md` — Bridge architecture, settings state model, side effects pattern

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `nativeEvent` WM_NCHITTEST handler in `window.py:460-492` — already handles resize edges/corners; extend with HTCAPTION for drag zones
- `scripts/verify-release.ps1` — existing verification script; expand with more checks rather than creating a new script
- `__main__.py:68` `main()` function — entry point where `--smoke-test` flag will be parsed
- `virelo/platform/resources.py` `resource_path()` — resolves paths for PyInstaller-compatible resource access (used in smoke test)

### Established Patterns
- PowerShell scripts use `$ErrorActionPreference = "Stop"` + `$LASTEXITCODE` checks (CLAUDE.md footgun #2)
- Version injection: build scripts parse APP_VERSION from config.py via regex
- `Virelo.spec` does NOT import virelo modules — uses regex for version (CLAUDE.md footgun #1)
- Bridge returns `{"ok": true, "data": ...}` / `{"ok": false, "error": "..."}` — not relevant to chrome work but context for smoke test

### Integration Points
- `window.py:147` — `FramelessWindowHint` already set; drag support is the missing complement
- `window.py:460-492` — `nativeEvent` is the existing WM_NCHITTEST handler; all hit-test modifications go here
- `.gitignore` — add `.planning/` entry
- `README.md` — expand with accurate product documentation

</code_context>

<specifics>
## Specific Ideas

- The signed lParam fix (D-04) is critical for multi-monitor setups where secondary monitors can have negative coordinates (e.g., monitor to the left of primary at -1920,0).
- UPX disabling (D-05) is a defensive measure — PySide6/Qt binaries can be fragile under UPX compression, causing intermittent crashes.
- The smoke test (D-06..D-08) enables CI to verify the app boots correctly without manual interaction, catching broken imports, missing resources, or initialization failures.
- The existing verify-release.ps1 already has the right structure — expanding it keeps the build pipeline consistent.

</specifics>

<deferred>
## Deferred Ideas

None — analysis stayed within phase scope.

</deferred>

---

*Phase: 08-window-chrome-and-release*
*Context gathered: 2026-04-24*
