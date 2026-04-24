# Requirements: Virelo

**Defined:** 2026-04-24
**Core Value:** The keyboard-triggered window snap must work reliably on any foreground window across all monitors, without interfering with fullscreen applications.

## v1.1 Requirements

Requirements for v1.1 Polish and Correctness. Each maps to roadmap phases.

### UI Action Placement

- [ ] **UI-01**: User sees Test snap only on the Window snap page (Target size card header) and in the command palette, not in the global footer
- [ ] **UI-02**: User can test snap with current visible draft values without saving first
- [ ] **UI-03**: User sees footer containing only status message, dirty indicator, Discard, and Save — no Reset defaults button
- [ ] **UI-04**: User can reset all settings to defaults from General page with a confirmation dialog
- [ ] **UI-05**: User can rebind snap and restore keys via key capture controls (press key to bind) instead of segmented SHIFT/CTRL/ALT controls
- [ ] **UI-06**: User sees Shortcuts page with accurate subtitle and no misleading rebind interaction hints
- [ ] **UI-07**: User sees only real, implemented actions in the command palette

### Bridge and Settings

- [ ] **BRDG-01**: User sees accurate dirty/clean state in the footer driven by Python dirty_changed signal, not local React inference
- [ ] **BRDG-02**: User can toggle Launch at login and have the startup shortcut created or removed on Save, with error reporting on failure
- [ ] **BRDG-03**: User can capture a new key binding and see it reflected as a dirty draft change before saving
- [ ] **BRDG-04**: User cannot cause silent boolean coercion bugs through bridge settings (strict parsing of true/false/1/0)
- [ ] **BRDG-05**: User can select System, Light, or Dark theme with Python sending both theme_mode and effective_theme to the frontend
- [ ] **BRDG-06**: User's UI preferences (accent, density, radius, sidebar mode, minimize-to-tray) are either persisted through the Python draft model or their controls are removed

### CI and Repo Hygiene

- [ ] **CI-01**: Developer can clone the repo and build without missing assets (icon.ico, branding/*, frontend/index.html committed)
- [ ] **CI-02**: Developer can run clean script to remove all generated artifacts recursively (__pycache__, *.pyc, .pytest_cache, .ruff_cache, frontend/dist, dist, build, installer/dist)
- [ ] **CI-03**: CI stale-name check passes without false positives on its own workflow file or test docstrings
- [ ] **CI-04**: CI unit tests pass (either on windows-latest or with platform-guarded Win32 imports on Ubuntu)
- [ ] **CI-05**: Frontend package.json version is synchronized with APP_VERSION from config.py
- [ ] **CI-06**: README states correct Python version requirement matching pyproject.toml (3.12+)
- [ ] **CI-07**: Installer support URL and config.py support URL use the same source of truth
- [ ] **CI-08**: Stale class names renamed to reflect configurable keys (ShiftSnapRestore → SnapRestoreController, HotkeyListener → MultiPressHotkeyListener)

### Window Chrome and Release

- [ ] **CHRM-01**: User can drag the frameless window from non-interactive title bar areas via Python-side event filter
- [ ] **CHRM-02**: User can resize the window correctly on monitors with negative screen coordinates (signed 16-bit lParam decoding)
- [ ] **CHRM-03**: Developer builds without UPX compression until release pipeline is verified stable
- [ ] **CHRM-04**: Developer can run a non-interactive smoke test (--smoke-test) that verifies resource paths, frontend dist, QWebEngine, settings, and bridge init
- [ ] **CHRM-05**: Developer can run release verification that checks version consistency, asset existence, and built content correctness
- [ ] **CHRM-06**: .planning/ is gitignored and not published; public docs exist in README.md and docs/
- [ ] **CHRM-07**: Public documentation covers build instructions, troubleshooting, and release checklist with correct product name and accurate descriptions

## Future Requirements

Deferred to v1.2+. Tracked but not in current roadmap.

### Module System

- **MOD-01**: Prepare a module registry architecture for future personal tools
- **MOD-02**: Preset snap sizes with keyboard cycling
- **MOD-03**: Per-process snap exclusion list

### Explorer Enhancements

- **EXPL-01**: Remember Explorer column widths per folder

## Out of Scope

| Feature | Reason |
|---------|--------|
| Plugin system or third-party extensibility | Personal-use tool, internal registry is sufficient |
| Cross-platform support | Windows-only by design, deep Win32/COM dependency |
| User authentication or multi-user features | Single-user desktop utility |
| Auto-update mechanism | Not implemented, do not add fake UI for it |
| Telemetry or analytics | Personal tool, do not add fake UI for it |
| Mobile or web deployment | Desktop only |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| UI-01 | — | Pending |
| UI-02 | — | Pending |
| UI-03 | — | Pending |
| UI-04 | — | Pending |
| UI-05 | — | Pending |
| UI-06 | — | Pending |
| UI-07 | — | Pending |
| BRDG-01 | — | Pending |
| BRDG-02 | — | Pending |
| BRDG-03 | — | Pending |
| BRDG-04 | — | Pending |
| BRDG-05 | — | Pending |
| BRDG-06 | — | Pending |
| CI-01 | — | Pending |
| CI-02 | — | Pending |
| CI-03 | — | Pending |
| CI-04 | — | Pending |
| CI-05 | — | Pending |
| CI-06 | — | Pending |
| CI-07 | — | Pending |
| CI-08 | — | Pending |
| CHRM-01 | — | Pending |
| CHRM-02 | — | Pending |
| CHRM-03 | — | Pending |
| CHRM-04 | — | Pending |
| CHRM-05 | — | Pending |
| CHRM-06 | — | Pending |
| CHRM-07 | — | Pending |

**Coverage:**
- v1.1 requirements: 28 total
- Mapped to phases: 0
- Unmapped: 28 ⚠️

---
*Requirements defined: 2026-04-24*
*Last updated: 2026-04-24 after initial definition*
