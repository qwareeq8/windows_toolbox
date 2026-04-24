# Phase 3: Structure and Quality - Context

**Gathered:** 2026-04-24
**Status:** Ready for planning

<domain>
## Phase Boundary

Reorganize the flat Python module layout into a `virelo/` package with six subpackages (app, bridge, services, workers, platform, settings), add quality tooling (Ruff lint/format, pytest, Vitest), and establish GitHub Actions CI that catches regressions on every push. The frontend is not restructured — only backend tests are added for it via Vitest.

</domain>

<decisions>
## Implementation Decisions

### Package layout
- **D-01:** Create `virelo/` package with subpackages: `app/`, `bridge/`, `services/`, `workers/`, `platform/`, `settings/` — matching STRUCT-01 exactly.
- **D-02:** File mapping to subpackages:
  - `app/` — `__main__.py` (startup only), `window.py` (MainWindow), `webview.py` (VireloWebView/VireloWebPage), `config.py` (app_config constants)
  - `bridge/` — `bridge.py` (VireloBridge), `capture_guard.py` (CaptureGuard)
  - `services/` — `snap.py` (SnapService), `explorer_columns.py` (COM IColumnManager interface)
  - `workers/` — `key_capture.py` (KeyCaptureWorker, KeyCaptureSession), `explorer.py` (ExplorerAutosizeEngine, ExplorerAutosizeWorker, TabAutosizeState)
  - `platform/` — `win32_helpers.py` (DPI, monitor rect, fullscreen detection, window geometry), `startup.py` (startup shortcut), `theme.py` (Windows registry theme detection)
  - `settings/` — `persistence.py` (Settings QSettings class), `state.py` (SettingsState JSON facade)
- **D-03:** Top-level `main.py` becomes a thin shim: `from virelo.app import main; main()` — satisfying STRUCT-02.
- **D-04:** `explorer_columns.py` stays as a single module despite exceeding 500 lines — justified under STRUCT-03's "unless justified" clause because it's a cohesive COM interface implementation where splitting would break the abstraction.

### Module splitting strategy
- **D-05:** Split `main.py` bottom-up to avoid circular imports and signal/slot disconnection (per STATE.md flag):
  1. Pure utility functions first (`resource_path`, DPI awareness, monitor rect helpers → `platform/win32_helpers.py`)
  2. ShiftSnapRestore engine next (self-contained snap logic → `services/snap_engine.py` or kept inside `services/snap.py`)
  3. MainWindow last (depends on everything else → `app/window.py`)
- **D-06:** Split `workers.py` (978 lines) into `workers/key_capture.py` and `workers/explorer.py` — each under 500 lines.
- **D-07:** COM apartment threading constraint: `ExplorerAutosizeWorker` COM initialization must stay co-located in the same file/class as its COM operations. Do not separate COM init from COM usage across modules.

### Test scope and tiering
- **D-08:** Two test tiers:
  - `tests/unit/` — Pure logic tests that run without Qt, PySide6, or admin elevation: snap geometry calculations, config defaults, settings key validation, theme mode normalization
  - `tests/integration/` — Tests requiring PySide6: bridge payload structure, settings state operations. Marked with `@pytest.mark.requires_qt` and excluded from CI.
- **D-09:** Frontend tests via Vitest: `bridgeToState`/`stateToBridge` mapping correctness, command palette filtering logic, component render smoke tests. No bridge communication tests (that's integration).
- **D-10:** Admin elevation is not available in GitHub Actions — all CI tests must run without admin privileges. Tests requiring admin are integration-tier and run locally only.

### Linting configuration
- **D-11:** Ruff configured in `pyproject.toml` with rules: E (pycodestyle errors), F (pyflakes), I (isort), UP (pyupgrade). Line length 100 (matches existing de facto style).
- **D-12:** Ruff formatter enabled (replaces need for Black). Format check enforced in CI.
- **D-13:** No frontend linter added in Phase 3. Vitest provides quality assurance for the frontend. A frontend linter can be added in a future phase if needed.

### CI pipeline
- **D-14:** Single `.github/workflows/ci.yml` workflow with four jobs:
  1. `lint` — `ruff check .` and `ruff format --check .`
  2. `test` — `pytest tests/unit/` (pure logic only, no Qt)
  3. `frontend` — `npm ci` + `npm run build` + `npx vitest run`
  4. `stale-name` — `grep -rn "Windows Toolbox" . --include="*.py" --include="*.jsx" --include="*.js"` fails if matches found
- **D-15:** Ubuntu runner for all jobs (fast, free-tier friendly). Unit tests are designed to be platform-independent. Tests requiring Windows/PySide6 are excluded from CI via pytest marker.
- **D-16:** Triggers: `push` and `pull_request` to `main` branch.
- **D-17:** Cache `pip` and `npm` dependencies between runs for speed.

### Claude's Discretion
- Exact Ruff rule exceptions for existing code that would be too noisy to fix in this phase
- pytest fixture organization and conftest.py structure
- Vitest configuration details (test file naming, setup files)
- Whether to add `py.typed` marker file
- pyproject.toml project metadata details beyond what's required for tooling

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Project vision, constraints, out-of-scope features
- `.planning/REQUIREMENTS.md` — STRUCT-01..06, QUAL-01..06 requirement definitions
- `CLAUDE.md` — Forbidden changes, known footguns (especially `app_config.py` import restriction in spec file)

### Prior phase context
- `.planning/phases/01-hygiene-and-build-pipeline/01-CONTEXT.md` — Build pipeline decisions, version consolidation
- `.planning/phases/02-security-and-bridge-hardening/02-CONTEXT.md` — Bridge architecture, draft model, WebEngine security

### Architecture
- `.planning/codebase/ARCHITECTURE.md` — Current architecture layers and data flow
- `.planning/codebase/STRUCTURE.md` — Current file layout (what gets reorganized)
- `.planning/codebase/CONVENTIONS.md` — Naming, imports, error handling patterns to preserve
- `.planning/codebase/CONCERNS.md` — Known concerns including main.py monolith

### Research
- `.planning/research/STACK.md` — Recommended tooling (Ruff, pytest, pyproject.toml)
- `.planning/research/ARCHITECTURE.md` — Package structure recommendations

### Current implementation (files being reorganized)
- `main.py` — 1449 lines, contains MainWindow + ShiftSnapRestore + utilities (primary split target)
- `workers.py` — 978 lines, contains key capture + explorer autosize (split target)
- `bridge.py` — 281 lines, VireloBridge (moves to `virelo/bridge/`)
- `settings.py` + `settings_state.py` — Settings persistence and state (move to `virelo/settings/`)
- `Virelo.spec` — PyInstaller spec (must be updated for new package paths)
- `frontend/vite.config.js` — May need `define` updates if config import path changes

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `app_config.py` — Constants and defaults, moves to `virelo/app/config.py` with same interface
- `snap_service.py` — Already a clean facade, moves to `virelo/services/snap.py`
- `capture_guard.py` — Thread-safe mutex, moves to `virelo/bridge/capture_guard.py`
- `startup_shortcut.py` — Standalone utility, moves to `virelo/platform/startup.py`
- `theme.py` — Standalone with injectable `read_registry` param, moves to `virelo/platform/theme.py`

### Established Patterns
- Dependency injection already used in `theme.py` (`read_registry=None`) and `startup_shortcut.py` (`exists=os.path.exists`) — extend this pattern to make more code testable
- `_safe_int`/`_safe_bool` in `settings.py` — defensive coercion pattern to preserve
- `SettingsState.KEYS` dict with type coercers — validation pattern to preserve and test
- JSON bridge protocol (`{ok, data}` / `{ok, error}`) — test this contract

### Integration Points
- `Virelo.spec` `datas=` and entry point — must update for `virelo/` package
- `scripts/build-app.ps1` — PyInstaller invocation may need path updates
- `frontend/src/bridge.js` — QWebChannel registration name `"bridge"` must not change
- `main.py` entry point — must remain at root for backward compatibility with existing shortcuts/scripts

</integration_points>

</code_context>

<specifics>
## Specific Ideas

- The package split should be mechanical — same code, new locations, updated imports. No refactoring of logic within this phase.
- `pyproject.toml` should define project metadata even though full Python packaging isn't needed yet (D-01 from Phase 1 context — sets up for quality tooling).
- Keep `main.py` at root as a thin entry point shim — don't force users/scripts to change `python main.py` to `python -m virelo`.
- Tests should focus on the four areas explicitly named in QUAL-03: settings validation, theme resolution, snap geometry, bridge payloads.

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 03-structure-and-quality*
*Context gathered: 2026-04-24*
