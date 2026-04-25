# Virelo

## What This Is

Virelo is a personal Windows desktop utility that snaps the foreground window to a configurable size and position with a multi-press keyboard shortcut, and auto-sizes File Explorer Detail view columns when navigating folders. It uses a Python/PySide6 backend hosting a React frontend inside QWebEngineView, communicating through QWebChannel.

## Core Value

The keyboard-triggered window snap must work reliably on any foreground window across all monitors, without interfering with fullscreen applications.

## Current Milestone: v1.1 Polish and Correctness

**Goal:** Make every visible UI control contextual, real, and backed by the same Python state model — plus fix CI, repo hygiene, and window chrome.

**Target features:**
- UI action placement: move Test snap to Window snap page, test current draft, wire dead buttons
- Bridge correctness: Python-owned dirty state, launch-at-login side effects, key capture via draft, strict booleans, coherent theme
- CI and repo hygiene: commit required assets, fix stale-name CI, fix Ubuntu CI, sync versions
- Window chrome and release: title-bar dragging, signed hit testing, smoke-test, release verification

## Requirements

### Validated

- ✓ Multi-press keyboard snap resizes and centers the foreground window — v1.0
- ✓ Keyboard-triggered restore returns the window to its original size and position — v1.0
- ✓ Game mode skips snapping when a fullscreen application is detected — v1.0
- ✓ Explorer column auto-size adjusts Detail view columns on folder navigation — v1.0
- ✓ System tray icon with show/quit actions — v1.0
- ✓ Settings persisted to Windows registry via QSettings — v1.0
- ✓ Single-instance mutex prevents duplicate launches — v1.0
- ✓ Admin elevation enforced at startup — v1.0
- ✓ React frontend renders all settings UI inside QWebEngineView — v1.0
- ✓ Dark/light/system theme with accent colors — v1.0
- ✓ Python-owned dirty state with dirty_changed signal — Phase 5
- ✓ Strict boolean parsing at bridge boundary (_strict_bool) — Phase 5
- ✓ All settings route through draft/commit model (no standalone slots) — Phase 5
- ✓ Theme System/Light/Dark with mode+effective distinction — Phase 5
- ✓ Accent, density, minimize-to-tray persisted through Python draft model — Phase 5
- ✓ Key capture produces draft change, not immediate persist — Phase 5
- ✓ Command palette for quick actions — v1.0
- ✓ Key capture workflow for rebinding snap and restore keys — v1.0
- ✓ All stale naming removed, version consolidated to single source of truth — v1.0
- ✓ Reproducible build pipeline (bootstrap through installer) — v1.0
- ✓ WebEngine security lockdown with draft/commit state model — v1.0
- ✓ Fake controls removed, all UI wired to Python bridge — v1.0
- ✓ Codebase restructured as testable virelo/ package with CI — v1.0
- ✓ Hotkey detection separated from window movement, ExplorerService extracted — v1.0
- ✓ Source assets committed for reproducible clone-and-build — Phase 7
- ✓ Clean script removes all generated artifacts recursively — Phase 7
- ✓ CI passes without false positives (stale-name, platform guards) — Phase 7
- ✓ Version strings synchronized across config.py, package.json, README, installer — Phase 7
- ✓ Stale class names renamed (SnapRestoreController, MultiPressHotkeyListener) — Phase 7

### Active

See `.planning/REQUIREMENTS.md` for v1.1 scoped requirements.

### Future

- Prepare a module registry architecture for future personal tools
- Preset snap sizes with keyboard cycling
- Per-process snap exclusion list
- Remember Explorer column widths per folder

### Out of Scope

- Plugin system or third-party extensibility — personal-use tool, internal registry is sufficient
- Cross-platform support — Windows-only by design, deep Win32/COM dependency
- User authentication or multi-user features — single-user desktop utility
- Auto-update mechanism — not implemented, do not add fake UI for it
- Telemetry or analytics — personal tool, do not add fake UI for it
- Mobile or web deployment — desktop only

## Context

Shipped v1.0 with 4,355 LOC Python, 1,482 LOC React/JS, 607 LOC tests. All "Windows Toolbox" naming removed. Codebase organized as a `virelo/` package with 6 subpackages (app, bridge, services, workers, platform, settings). 53 unit tests pass, Ruff linting clean, GitHub Actions CI enforces on every push. Reproducible build pipeline produces installer from clean checkout. WebEngine locked down, bridge uses draft/commit state model, all UI controls wired to Python backend. Snap and Explorer features are isolated services with proper separation of concerns.

## Constraints

- **Platform**: Windows 10/11 x64 only — enforced at startup
- **Privileges**: Requires administrator elevation — UAC prompt at launch
- **Stack**: Python 3 + PySide6 backend, React 19 + Vite frontend, QWebChannel bridge — established and not changing
- **Packaging**: PyInstaller for app bundle, Inno Setup for installer — established
- **Scope**: Personal-use software — optimize for developer productivity, not enterprise concerns

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Keep "Virelo" as product name | Abstract enough for future personal utilities, not limited to snapping | — Pending |
| React inside QWebEngineView for UI | Already migrated, correct direction for design fidelity | ✓ Good |
| QWebChannel for Python-JS bridge | Qt-native, documented, works with WebEngine | ✓ Good |
| Do not commit frontend/dist | Build artifact, regenerated by pipeline | — Pending |
| Coarse phase granularity | Fewer broader phases for this cleanup/hardening work | — Pending |
| Quality model profile (Opus) for planning agents | Deeper analysis for architectural cleanup decisions | — Pending |
| _strict_bool replaces bool for bridge booleans | Prevents bool("false")==True coercion bug | ✓ Good |
| Accent/density validation via tuple allowlist | Matches project style, no enum class overhead | ✓ Good |
| dirty_changed signal replaces manual setUnsaved | Single source of truth for dirty state | ✓ Good |
| Remove apply_theme/toggle_run_at_startup slots | All settings through draft/commit, no bypasses | ✓ Good |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `/gsd-transition`):
1. Requirements invalidated? → Move to Out of Scope with reason
2. Requirements validated? → Move to Validated with phase reference
3. New requirements emerged? → Add to Active
4. Decisions to log? → Add to Key Decisions
5. "What This Is" still accurate? → Update if drifted

**After each milestone** (via `/gsd-complete-milestone`):
1. Full review of all sections
2. Core Value check — still the right priority?
3. Audit Out of Scope — reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-04-25 after Phase 7*
