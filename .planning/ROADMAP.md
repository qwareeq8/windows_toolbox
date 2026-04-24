# Roadmap: Virelo

## Overview

Virelo is a working Windows desktop utility that needs hardening and cleanup before the repository is presentable. The roadmap progresses from low-risk hygiene (naming, repo files, build pipeline) through security and bridge hardening, into structural refactoring with quality infrastructure, and finishes with feature-specific improvements to the snap and Explorer services. Each phase delivers a verifiable capability increment, ordered so that earlier phases unblock later ones and regressions are caught by the quality infrastructure established in Phase 3.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [x] **Phase 1: Hygiene and Build Pipeline** - Remove stale naming, create repo files, and establish a reproducible build from clean checkout
- [ ] **Phase 2: Security and Bridge Hardening** - Lock down WebEngine, restructure the bridge with draft state, and wire all frontend controls
- [ ] **Phase 3: Structure and Quality** - Add quality tooling, split main.py into a virelo package, and establish CI
- [ ] **Phase 4: Snap and Explorer Hardening** - Extract and harden the snap and Explorer services with tests and correct scope

## Phase Details

### Phase 1: Hygiene and Build Pipeline
**Goal**: A clean checkout produces a correctly-named, installable application with no stale references, and the repository has standard open-source files
**Depends on**: Nothing (first phase)
**Requirements**: IDENT-01, IDENT-02, IDENT-03, IDENT-04, REPO-01, REPO-02, REPO-03, REPO-04, REPO-05, BUILD-01, BUILD-02, BUILD-03, BUILD-04, BUILD-05, BUILD-06
**Success Criteria** (what must be TRUE):
  1. Running the build from a clean checkout (no prior node_modules, no prior dist, no prior .venv) produces a launchable Virelo.exe that loads the React frontend offline
  2. No file in the repository contains "Windows Toolbox" (verified by grep)
  3. Version string is defined once in app_config.py and appears correctly in the installer, the frontend About view, and pyproject.toml without manual duplication
  4. Repository root contains .gitignore, README, LICENSE (MIT), and CLAUDE.md, and git status shows no untracked generated artifacts
  5. No deprecated Qt attribute warnings appear in the application log at startup
**Plans**: 3 plans

Plans:
- [x] 01-01-PLAN.md -- Identity cleanup, version consolidation, and stale reference removal
- [x] 01-02-PLAN.md -- Repository files (.gitignore, README, LICENSE, CLAUDE.md)
- [x] 01-03-PLAN.md -- Build pipeline (6 PowerShell scripts)

### Phase 2: Security and Bridge Hardening
**Goal**: The WebEngine host is locked down for an admin-elevated process, the bridge uses structured draft/commit state management, and every visible frontend control connects to the Python backend
**Depends on**: Phase 1
**Requirements**: SEC-01, SEC-02, SEC-03, SEC-04, SEC-05, BRDG-01, BRDG-02, BRDG-03, BRDG-04, BRDG-05, BRDG-06, UI-01, UI-02, UI-03, UI-04, UI-05, UI-06
**Success Criteria** (what must be TRUE):
  1. Attempting to navigate the WebEngine to an external URL (e.g., https://example.com) is blocked and does not load -- verified by injecting a navigation request
  2. Changing a setting in the frontend creates a draft; clicking Save persists it to QSettings and applies side effects (e.g., startup shortcut), clicking Discard reverts to persisted values
  3. No fake controls are visible in the UI -- auto-update, telemetry, hidden files, file extensions, and remember-columns toggles are gone
  4. Every visible toggle, slider, button, and command palette action either calls a real bridge method or has been removed
  5. The bridge returns structured payloads ({ok, data} or {ok, error}) for all operations, and unknown keys produce a structured error response
**Plans**: TBD
**UI hint**: yes

Plans:
- [ ] 02-01: TBD
- [ ] 02-02: TBD
- [ ] 02-03: TBD

### Phase 3: Structure and Quality
**Goal**: The codebase is organized as a testable virelo/ package with linting, tests, and CI catching regressions on every push
**Depends on**: Phase 2
**Requirements**: STRUCT-01, STRUCT-02, STRUCT-03, STRUCT-04, STRUCT-05, STRUCT-06, QUAL-01, QUAL-02, QUAL-03, QUAL-04, QUAL-05, QUAL-06
**Success Criteria** (what must be TRUE):
  1. Python source lives under a virelo/ package with subpackages (app, bridge, services, workers, platform, settings) and main.py or __main__.py contains only startup code
  2. Ruff lint and format pass with zero errors on the entire Python codebase, and CI enforces this on every pull request
  3. pytest runs with passing tests for settings validation, theme resolution, snap geometry calculations, and bridge payload structure -- without launching the full UI
  4. Frontend tests run via Vitest for key component behaviors
  5. CI fails if "Windows Toolbox" reappears in any source file (stale-name regression gate)
**Plans**: TBD

Plans:
- [ ] 03-01: TBD
- [ ] 03-02: TBD
- [ ] 03-03: TBD

### Phase 4: Snap and Explorer Hardening
**Goal**: The snap and Explorer features operate as isolated, tested services with correct scope and robust edge-case handling
**Depends on**: Phase 3
**Requirements**: SNAP-01, SNAP-02, SNAP-03, SNAP-04, SNAP-05, EXPL-01, EXPL-02, EXPL-03
**Success Criteria** (what must be TRUE):
  1. Snap logic lives in a dedicated service module with hotkey detection separated from window movement, and Virelo's own window is excluded from snapping
  2. Restoring a previously-maximized window returns it to maximized state, not the pre-maximized geometry
  3. Snap geometry calculations have unit tests covering at least: single monitor, two-monitor horizontal layout, and a monitor with negative coordinates
  4. Explorer settings page shows only the auto-size columns feature (no hidden-files, file-extensions, or remember-columns controls)
  5. Explorer worker starts only when its setting is enabled, stops cleanly on quit, and its orchestration logic lives in a service module rather than MainWindow
**Plans**: TBD

Plans:
- [ ] 04-01: TBD
- [ ] 04-02: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Hygiene and Build Pipeline | 3/3 | Complete   | 2026-04-24 |
| 2. Security and Bridge Hardening | 0/3 | Not started | - |
| 3. Structure and Quality | 0/3 | Not started | - |
| 4. Snap and Explorer Hardening | 0/2 | Not started | - |
