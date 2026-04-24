# Requirements: Virelo

**Defined:** 2026-04-24
**Core Value:** The keyboard-triggered window snap must work reliably on any foreground window across all monitors, without interfering with fullscreen applications.

## v1 Requirements

Requirements for the hardening and cleanup milestone. Each maps to roadmap phases.

### Identity

- [x] **IDENT-01**: All "Windows Toolbox" references removed from source, spec files, and build scripts
- [x] **IDENT-02**: All product metadata (name, version, ID, publisher) defined in one source of truth in app_config.py
- [x] **IDENT-03**: Version string flows from app_config.py to installer, frontend, and package metadata without manual duplication
- [x] **IDENT-04**: Internal migration-phase comments (Phase 7, SC-6, IC-11, v2) removed from production source

### Repository

- [x] **REPO-01**: .gitignore excludes all generated artifacts (.venv, __pycache__, build, dist, frontend/node_modules, frontend/dist, logs)
- [x] **REPO-02**: README explains what Virelo does, Windows-only requirement, admin privileges, build from source, and current status
- [x] **REPO-03**: LICENSE file present (MIT)
- [x] **REPO-04**: CLAUDE.md provides Claude Code with build commands, conventions, and project layout
- [x] **REPO-05**: Deprecated Qt attributes (AA_EnableHighDpiScaling, AA_UseHighDpiPixmaps) removed

### Build

- [x] **BUILD-01**: Clean checkout produces working app with one command sequence (bootstrap, build frontend, PyInstaller, installer)
- [x] **BUILD-02**: Build fails early if npm, node, Python, PyInstaller, or ISCC is missing
- [x] **BUILD-03**: Build fails early if frontend/dist is absent after frontend build step
- [x] **BUILD-04**: PyInstaller spec renamed from "Windows Toolbox.spec" to "Virelo.spec"
- [x] **BUILD-05**: Inno Setup reads version from generated metadata, not hardcoded string
- [x] **BUILD-06**: Installed app launches and loads React frontend offline

### Security

- [x] **SEC-01**: WebEngine blocks external navigation via acceptNavigationRequest override
- [x] **SEC-02**: LocalContentCanAccessRemoteUrls disabled in release mode
- [x] **SEC-03**: Dev mode triggered only by VIRELO_DEV=1 environment variable, not sys.frozen check
- [x] **SEC-04**: Missing frontend build shows clear error message, not blank page
- [x] **SEC-05**: Default WebEngine context menu disabled in release mode

### Bridge

- [x] **BRDG-01**: Python-side draft model holds unsaved changes separately from persisted settings
- [x] **BRDG-02**: Save commits draft to QSettings and applies side effects (startup shortcut, snap manager, Explorer worker)
- [x] **BRDG-03**: Discard reverts draft to persisted settings and reapplies runtime bindings
- [x] **BRDG-04**: Unknown bridge keys rejected with structured error, not silently dropped
- [x] **BRDG-05**: Bridge returns structured payloads ({ok, data} or {ok, error}) for all operations
- [x] **BRDG-06**: Launch-at-login toggle creates or removes startup shortcut on save

### Frontend

- [ ] **UI-01**: Fake controls removed (auto-update, telemetry, hidden files, file extensions, remember columns)
- [ ] **UI-02**: No-op command palette actions removed or wired to real bridge calls
- [ ] **UI-03**: Title bar minimize and close call bridge setWindowCommand and work in frameless window
- [ ] **UI-04**: Key capture uses bridge startKeyCapture/cancelKeyCapture, shows "Press a key" during capture
- [ ] **UI-05**: Every visible toggle, slider, and button connects to the Python bridge
- [ ] **UI-06**: React state survives save, discard, reset, and external Python-pushed updates

### Structure

- [ ] **STRUCT-01**: Python source organized as virelo/ package with subpackages (app, bridge, services, workers, platform, settings)
- [ ] **STRUCT-02**: main.py or __main__.py contains only app startup, not business logic
- [ ] **STRUCT-03**: No source file exceeds 500 lines unless justified
- [ ] **STRUCT-04**: Snap logic testable without launching the full UI
- [ ] **STRUCT-05**: Settings validation testable without WebEngine
- [ ] **STRUCT-06**: Duplicate code consolidated (path canonicalization, resource_path, autosize functions)

### Quality

- [ ] **QUAL-01**: pyproject.toml defines project metadata, dependencies, and tool configuration
- [ ] **QUAL-02**: Ruff configured for linting and formatting with rules enforced in CI
- [ ] **QUAL-03**: pytest configured with tests for settings validation, theme resolution, snap geometry, and bridge payloads
- [ ] **QUAL-04**: Frontend tests configured with Vitest for key component behaviors
- [ ] **QUAL-05**: GitHub Actions CI runs lint, tests, frontend build, and stale-name grep on pull requests
- [ ] **QUAL-06**: CI fails if "Windows Toolbox" reappears in any source file

### Snap

- [ ] **SNAP-01**: Snap logic extracted from main.py into a dedicated service module
- [ ] **SNAP-02**: Hotkey detection separated from window movement logic
- [ ] **SNAP-03**: Virelo's own window excluded from snapping
- [ ] **SNAP-04**: Restore correctly handles previously-maximized windows
- [ ] **SNAP-05**: Geometry calculations have unit tests covering multi-monitor scenarios

### Explorer

- [ ] **EXPL-01**: Explorer page shows only implemented features (auto-size columns)
- [ ] **EXPL-02**: Explorer worker orchestration moved out of MainWindow into a service
- [ ] **EXPL-03**: Explorer worker starts only when the setting is enabled and stops cleanly on quit

## v2 Requirements

Deferred to future milestone. Tracked but not in current roadmap.

### Explorer Enhancements

- **EXPL-10**: Remember column widths per folder
- **EXPL-11**: Show/hide hidden files toggle
- **EXPL-12**: Show/hide file extensions toggle

### Snap Enhancements

- **SNAP-10**: Preset snap sizes (50%, 66%, 76%, 90%)
- **SNAP-11**: Cycle preset sizes via keyboard shortcut
- **SNAP-12**: Per-process exclusion list
- **SNAP-13**: Snap diagnostics command ("Why did snap not run?")

### Module Architecture

- **ARCH-01**: Internal module registry with ModuleInfo dataclass
- **ARCH-02**: Snap and Explorer registered as modules
- **ARCH-03**: Adding a new feature does not require editing five unrelated files

### Future Tools

- **TOOL-01**: Window layout presets
- **TOOL-02**: Monitor profile presets
- **TOOL-03**: Clipboard utilities
- **TOOL-04**: Diagnostics and logs viewer

## Out of Scope

Explicitly excluded. Documented to prevent scope creep.

| Feature | Reason |
|---------|--------|
| Plugin system / third-party extensibility | Personal-use tool, internal registry sufficient |
| Cross-platform support | Deep Win32/COM dependency, porting is a rewrite |
| Auto-update mechanism | No update server, no code signing, no infrastructure |
| Telemetry / analytics | Personal tool, one user, no benefit |
| Code signing certificate | $200-400/year cost for personal tool with one user |
| MSI installer | Inno Setup works and is already configured |
| Documentation website | README + CLAUDE.md sufficient for personal tool |
| Internationalization | Single-user, English-only tool |
| Custom crash reporting service | Local crash.log and virelo.log sufficient |
| Mobile or web deployment | Desktop-only by design |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| IDENT-01 | Phase 1 | Complete |
| IDENT-02 | Phase 1 | Complete |
| IDENT-03 | Phase 1 | Complete |
| IDENT-04 | Phase 1 | Complete |
| REPO-01 | Phase 1 | Complete |
| REPO-02 | Phase 1 | Complete |
| REPO-03 | Phase 1 | Complete |
| REPO-04 | Phase 1 | Complete |
| REPO-05 | Phase 1 | Complete |
| BUILD-01 | Phase 1 | Complete |
| BUILD-02 | Phase 1 | Complete |
| BUILD-03 | Phase 1 | Complete |
| BUILD-04 | Phase 1 | Complete |
| BUILD-05 | Phase 1 | Complete |
| BUILD-06 | Phase 1 | Complete |
| SEC-01 | Phase 2 | Complete |
| SEC-02 | Phase 2 | Complete |
| SEC-03 | Phase 2 | Complete |
| SEC-04 | Phase 2 | Complete |
| SEC-05 | Phase 2 | Complete |
| BRDG-01 | Phase 2 | Complete |
| BRDG-02 | Phase 2 | Complete |
| BRDG-03 | Phase 2 | Complete |
| BRDG-04 | Phase 2 | Complete |
| BRDG-05 | Phase 2 | Complete |
| BRDG-06 | Phase 2 | Complete |
| UI-01 | Phase 2 | Pending |
| UI-02 | Phase 2 | Pending |
| UI-03 | Phase 2 | Pending |
| UI-04 | Phase 2 | Pending |
| UI-05 | Phase 2 | Pending |
| UI-06 | Phase 2 | Pending |
| STRUCT-01 | Phase 3 | Pending |
| STRUCT-02 | Phase 3 | Pending |
| STRUCT-03 | Phase 3 | Pending |
| STRUCT-04 | Phase 3 | Pending |
| STRUCT-05 | Phase 3 | Pending |
| STRUCT-06 | Phase 3 | Pending |
| QUAL-01 | Phase 3 | Pending |
| QUAL-02 | Phase 3 | Pending |
| QUAL-03 | Phase 3 | Pending |
| QUAL-04 | Phase 3 | Pending |
| QUAL-05 | Phase 3 | Pending |
| QUAL-06 | Phase 3 | Pending |
| SNAP-01 | Phase 4 | Pending |
| SNAP-02 | Phase 4 | Pending |
| SNAP-03 | Phase 4 | Pending |
| SNAP-04 | Phase 4 | Pending |
| SNAP-05 | Phase 4 | Pending |
| EXPL-01 | Phase 4 | Pending |
| EXPL-02 | Phase 4 | Pending |
| EXPL-03 | Phase 4 | Pending |

**Coverage:**
- v1 requirements: 52 total
- Mapped to phases: 52
- Unmapped: 0

---
*Requirements defined: 2026-04-24*
*Last updated: 2026-04-24 after roadmap creation*
