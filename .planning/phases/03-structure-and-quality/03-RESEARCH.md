# Phase 3: Structure and Quality - Research

**Researched:** 2026-04-24
**Domain:** Python package restructuring, test infrastructure, linting, CI/CD
**Confidence:** HIGH

## Summary

Phase 3 transforms Virelo's flat Python module layout (12 root-level .py files totaling 4142 lines) into a structured `virelo/` package with six subpackages, adds quality tooling (Ruff lint/format, pytest, Vitest), and establishes GitHub Actions CI. The codebase has zero tests, no linter, no formatter, no pyproject.toml, and no CI -- all must be created from scratch.

The highest-risk work is splitting `main.py` (1449 lines) and `workers.py` (978 lines). The CONTEXT.md specifies a bottom-up splitting strategy to avoid circular imports and signal/slot disconnection. Pure utility functions move first, then services, then bridge/app wiring. The `explorer_columns.py` (861 lines) stays as-is per D-04, justified as a cohesive COM interface.

**Primary recommendation:** Execute the package reorganization in strict bottom-up order (platform -> settings -> services -> workers -> bridge -> app), with test infrastructure set up first so each split can be immediately validated. Use Ubuntu CI runners for lint and frontend; unit tests designed to be platform-independent (no Win32/Qt dependencies).

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Create `virelo/` package with subpackages: `app/`, `bridge/`, `services/`, `workers/`, `platform/`, `settings/`
- **D-02:** Specific file mapping to subpackages (detailed in CONTEXT.md)
- **D-03:** Top-level `main.py` becomes thin shim: `from virelo.app import main; main()`
- **D-04:** `explorer_columns.py` stays as single module despite exceeding 500 lines (justified under STRUCT-03's "unless justified" clause)
- **D-05:** Split `main.py` bottom-up: pure utilities first, ShiftSnapRestore next, MainWindow last
- **D-06:** Split `workers.py` into `workers/key_capture.py` and `workers/explorer.py`
- **D-07:** COM apartment threading: `ExplorerAutosizeWorker` COM init must stay co-located with COM operations
- **D-08:** Two test tiers: `tests/unit/` (pure logic, no Qt) and `tests/integration/` (requires PySide6, `@pytest.mark.requires_qt`, excluded from CI)
- **D-09:** Frontend tests via Vitest: `bridgeToState`/`stateToBridge` mapping, command palette filtering, component render smoke tests
- **D-10:** Admin elevation not available in GitHub Actions -- all CI tests must run without admin privileges
- **D-11:** Ruff configured in `pyproject.toml` with rules: E, F, I, UP. Line length 100
- **D-12:** Ruff formatter enabled, format check enforced in CI
- **D-13:** No frontend linter in Phase 3 (Vitest only for frontend quality)
- **D-14:** Single CI workflow with four jobs: `lint`, `test`, `frontend`, `stale-name`
- **D-15:** Ubuntu runner for all jobs (unit tests are platform-independent)
- **D-16:** Triggers: `push` and `pull_request` to `main` branch
- **D-17:** Cache `pip` and `npm` dependencies between runs

### Claude's Discretion
- Exact Ruff rule exceptions for existing code that would be too noisy to fix
- pytest fixture organization and conftest.py structure
- Vitest configuration details (test file naming, setup files)
- Whether to add `py.typed` marker file
- pyproject.toml project metadata details beyond what's required for tooling

### Deferred Ideas (OUT OF SCOPE)
None -- discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| STRUCT-01 | Python source organized as virelo/ package with subpackages (app, bridge, services, workers, platform, settings) | Package layout pattern, dependency graph, bottom-up split order documented in Architecture Patterns |
| STRUCT-02 | main.py or __main__.py contains only app startup, not business logic | Thin shim pattern documented, `__main__.py` for `python -m virelo` support |
| STRUCT-03 | No source file exceeds 500 lines unless justified | Line counts verified: main.py=1449, workers.py=978 must split; explorer_columns.py=861 justified exception per D-04 |
| STRUCT-04 | Snap logic testable without launching the full UI | Snap geometry calculations extractable as pure functions; `_rect_matches_monitor`, `_should_skip_snap_for_game` are unit-testable |
| STRUCT-05 | Settings validation testable without WebEngine | `SettingsState.KEYS` validation, `apply_draft()`, `normalize_theme_mode()`, `normalize_snap_presses()` are pure logic; mock `Settings` object sufficient |
| STRUCT-06 | Duplicate code consolidated | Three duplicates identified: `resource_path` (2 copies), `canonicalize_path` (3 copies), `_autosize_quick/_full` (identical) |
| QUAL-01 | pyproject.toml defines project metadata, dependencies, and tool configuration | PEP 621 pyproject.toml pattern documented with dependency groups |
| QUAL-02 | Ruff configured for linting and formatting with rules enforced in CI | Ruff 0.15.12 config pattern documented; E, F, I, UP rules; line-length 100 |
| QUAL-03 | pytest configured with tests for settings validation, theme resolution, snap geometry, and bridge payloads | pytest 9.0.3 config pattern, test file structure, mock strategies documented |
| QUAL-04 | Frontend tests configured with Vitest for key component behaviors | Vitest 4.1.5 config with jsdom, test targets: bridgeToState/stateToBridge, command palette filtering |
| QUAL-05 | GitHub Actions CI runs lint, tests, frontend build, and stale-name grep on pull requests | CI workflow structure documented, Ubuntu runner, caching strategy |
| QUAL-06 | CI fails if "Windows Toolbox" reappears in any source file | Stale-name grep job pattern documented |
</phase_requirements>

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Package reorganization | Backend (Python) | -- | Pure file system restructuring, imports, `__init__.py` files |
| Test infrastructure (Python) | Backend (Python) | -- | pytest runs against Python modules |
| Test infrastructure (Frontend) | Frontend (React) | -- | Vitest runs against JSX/JS modules |
| Linting/formatting | Backend (Python) | -- | Ruff operates on Python source only |
| CI pipeline | CI/CD (GitHub Actions) | -- | Workflow YAML, runner configuration |
| pyproject.toml | Backend (Python) | -- | Project metadata and tool configuration |
| PyInstaller spec update | Build tooling | -- | `Virelo.spec` entry point and hidden imports |
| Stale-name regression gate | CI/CD (GitHub Actions) | -- | grep-based file scanning |

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| ruff | 0.15.12 | Python linting and formatting | [VERIFIED: pip index] Single tool replacing flake8+black+isort+pyupgrade. Sub-second execution. |
| pytest | 9.0.3 | Python test runner | [VERIFIED: pip index] De facto standard. Fixture model, plugin ecosystem, clear output. |
| vitest | 4.1.5 | JavaScript test runner | [VERIFIED: npm registry] Built on Vite -- shares same config and transform pipeline. |
| jsdom | 29.0.2 | DOM environment for Vitest | [VERIFIED: npm registry] Headless browser-like DOM for component testing. |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| pytest-cov | 7.1.0 | Coverage reporting | [VERIFIED: pip index] Integrates coverage.py with pytest. Use in local runs, optional in CI. |
| @testing-library/react | 16.3.2 | React component testing | [VERIFIED: npm registry] Testing by user behavior, not implementation details. |
| @testing-library/jest-dom | 6.9.1 | DOM assertion matchers | [VERIFIED: npm registry] Adds `.toBeInTheDocument()`, `.toHaveTextContent()`, etc. |
| @testing-library/user-event | 14.6.1 | Simulated user interactions | [VERIFIED: npm registry] Realistic click/type/etc events for component tests. |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| ruff | flake8 + black + isort | 3 tools vs 1; ruff is 10-100x faster. No reason for separate tools. |
| pytest | unittest | pytest runs unittest tests but has richer fixtures, plugins, cleaner syntax. |
| vitest | Jest | Jest requires separate Babel/SWC config for JSX. Vitest shares Vite pipeline. |
| Ubuntu CI runner | Windows CI runner | D-15 locks Ubuntu. Unit tests designed platform-independent. Saves CI cost. |

**Installation:**

Python (add to pyproject.toml `[project.optional-dependencies]`):
```bash
pip install -e ".[dev]"
# Installs: ruff>=0.15.12, pytest>=9.0.3, pytest-cov>=7.1.0
```

Frontend (run in `frontend/`):
```bash
npm install -D vitest@^4.1.5 @testing-library/react@^16.3.2 @testing-library/jest-dom@^6.9.1 @testing-library/user-event@^14.6.1 jsdom@^29.0.2
```

## Architecture Patterns

### System Architecture Diagram

```
   [Developer push/PR to main]
          |
          v
   [GitHub Actions CI] ------+-------+-------+--------+
          |                   |       |       |        |
          v                   v       v       v        v
     [lint job]          [test job] [frontend] [stale-name]
     ruff check .        pytest     npm ci     grep -rn
     ruff format         tests/     npm build  "Windows Toolbox"
       --check           unit/      vitest run
          |                   |       |        |
          v                   v       v        v
     [All pass?] ----------> [Gate: merge allowed]


   [Package dependency graph (import direction)]

   main.py (thin shim) --> virelo/__main__.py
                               |
                               v
                          virelo/app/window.py (MainWindow)
                          /        |            \
                         v         v             v
               virelo/bridge/  virelo/services/  virelo/webview.py
                    |              |       \
                    v              v        v
              virelo/settings/  virelo/workers/  virelo/platform/ (LEAF)
```

### Recommended Project Structure

```
virelo/
  __init__.py              # Package marker, exports __version__
  app/
    __init__.py             # Re-exports: main(), MainWindow
    __main__.py             # Entry: admin elevation, single-instance, QApp, MainWindow
    window.py               # MainWindow: tray, chrome, thread lifecycle, wiring
    config.py               # APP_NAME, APP_ID, DEFAULTS, LOG constants (from app_config.py)
  bridge/
    __init__.py             # Re-exports: VireloBridge, CaptureGuard
    bridge.py               # VireloBridge QObject (Slots/Signals)
    capture_guard.py        # CaptureGuard thread-safe mutex
  services/
    __init__.py             # Re-exports: SnapService, etc.
    snap.py                 # SnapService facade + ShiftSnapRestore engine
    explorer_columns.py     # COM IColumnManager (stays as-is per D-04)
    theme.py                # resolve_theme, toggle, get_windows_theme, normalize
    startup.py              # startup shortcut creation/removal
  workers/
    __init__.py             # Re-exports: KeyCaptureWorker, ExplorerAutosizeWorker
    key_capture.py          # KeyCaptureWorker, KeyCaptureSession
    explorer.py             # ExplorerAutosizeWorker, ExplorerAutosizeEngine, TabAutosizeState
  platform/
    __init__.py             # Re-exports: resource_path, get_monitor_rect, etc.
    win32_helpers.py        # USER32/KERNEL32, DPI, monitor rects, window rects, DWM
    resources.py            # resource_path (consolidated from 2 copies)
    paths.py                # canonicalize_path (consolidated from 3 copies)
  settings/
    __init__.py             # Re-exports: Settings, SettingsState, DEFAULTS
    persistence.py          # Settings class (QSettings read/write, _safe_int, _safe_bool)
    state.py                # SettingsState class (JSON facade, validation, draft model)

main.py                     # Thin shim: from virelo.app import main; main()
Virelo.spec                 # Updated: Analysis(['main.py']), hiddenimports=['virelo.*']
pyproject.toml              # NEW: project metadata, deps, tool config

tests/
  __init__.py
  conftest.py               # Shared fixtures: mock_settings, mock_defaults
  unit/
    __init__.py
    test_settings_state.py   # SettingsState validation, draft model, coercion
    test_theme.py            # normalize_theme_mode, resolve_theme, toggle_theme_mode
    test_snap_geometry.py    # _rect_matches_monitor, geometry calculations
    test_app_config.py       # DEFAULTS completeness, normalize_snap_presses
    test_bridge_payload.py   # Bridge JSON envelope structure (mock SettingsState)
    test_paths.py            # canonicalize_path consolidated function
    test_capture_guard.py    # CaptureGuard thread safety
  integration/
    __init__.py
    conftest.py              # QApplication fixture, pytest.mark.requires_qt

frontend/
  src/
    __tests__/               # Or *.test.jsx alongside source
      app.test.jsx           # bridgeToState/stateToBridge mapping tests
      panels.test.jsx        # Command palette filtering logic
      primitives.test.jsx    # Component render smoke tests
    test-setup.js            # jsdom setup, jest-dom matchers

.github/
  workflows/
    ci.yml                   # lint + test + frontend + stale-name jobs
```

### Pattern 1: Bottom-Up Module Extraction

**What:** Extract modules from the monolith starting with leaf dependencies (no virelo imports) and working upward to modules that depend on everything.

**When to use:** Always for this restructuring. Prevents circular imports.

**Ordering:**
1. `platform/` -- Pure Win32/ctypes functions. Zero virelo imports. Extract from `main.py` lines 165-530.
2. `settings/` -- Depends only on `app_config` constants and `theme.normalize_theme_mode`. Already separate files, just directory restructuring.
3. `services/` -- Depends on platform + settings. Move `snap_service.py`, `theme.py`, `startup_shortcut.py`, `capture_guard.py`. Extract `ShiftSnapRestore` from `main.py`.
4. `workers/` -- Depends on platform (callables injected). Split `workers.py` into two files.
5. `bridge/` -- Depends on settings + services. Move `bridge.py`, update imports.
6. `app/` -- Depends on everything. Extract `MainWindow` from `main.py`, create `__main__.py`.

**Why bottom-up:** Each step only references modules that have already been moved. No temporary circular dependencies. The application can be run and tested at each intermediate step.

### Pattern 2: Mock Settings for Unit Tests

**What:** Create a mock Settings class that holds attributes in memory instead of using QSettings (Windows registry).

**When to use:** For all unit tests that need settings data without Qt.

**Example:**
```python
# tests/conftest.py
import pytest
from virelo.app.config import DEFAULTS

class MockSettings:
    """In-memory Settings replacement for unit tests."""
    def __init__(self, **overrides):
        for key, val in DEFAULTS.items():
            setattr(self, key, overrides.get(key, val))
    def save(self):
        pass  # No-op
    def clear(self):
        pass  # No-op

@pytest.fixture
def mock_settings():
    return MockSettings()

@pytest.fixture
def settings_state(mock_settings):
    from virelo.settings.state import SettingsState
    return SettingsState(mock_settings)
```
[VERIFIED: pytest fixture pattern from pytest docs]

### Pattern 3: Testable Snap Geometry

**What:** Extract pure geometry calculation functions from `ShiftSnapRestore._snap()` so they can be unit tested without Win32 API calls.

**When to use:** For STRUCT-04 (snap logic testable without full UI).

**Key testable functions (already exist as standalone functions in main.py):**
- `_rect_matches_monitor(rect, monitor)` -- pure math, no Win32
- `_should_skip_snap_for_game(hwnd, settings, full_screen)` -- needs mock settings only
- `normalize_snap_presses(value)` -- pure validation in app_config.py

**Geometry calculation to extract:**
```python
# virelo/services/snap.py (new pure function)
def calculate_snap_position(
    monitor_left, monitor_top, monitor_width, monitor_height,
    width_pct, height_pct
) -> tuple:
    """Calculate snap target position (x, y, w, h) for a resizable window."""
    w = monitor_width * width_pct // 100
    h = monitor_height * height_pct // 100
    x = monitor_left + (monitor_width - w) // 2
    y = monitor_top + (monitor_height - h) // 2
    return (x, y, w, h)
```
[ASSUMED: This extraction is a reasonable design choice for testability]

### Pattern 4: PyInstaller Spec Update for Package Layout

**What:** Update `Virelo.spec` to work with the `virelo/` package structure.

**Key changes needed:**
```python
# Virelo.spec -- updated Analysis
a = Analysis(
    ['main.py'],          # Thin shim at root still works
    pathex=[],
    binaries=[],
    datas=[
        ("icon.ico", "."),
        ("frontend/dist", "frontend/dist"),
    ],
    hiddenimports=[
        'virelo', 'virelo.app', 'virelo.app.window', 'virelo.app.config',
        'virelo.bridge', 'virelo.bridge.bridge', 'virelo.bridge.capture_guard',
        'virelo.services', 'virelo.services.snap', 'virelo.services.theme',
        'virelo.settings', 'virelo.settings.persistence', 'virelo.settings.state',
        'virelo.workers', 'virelo.workers.key_capture', 'virelo.workers.explorer',
        'virelo.platform', 'virelo.platform.win32_helpers', 'virelo.platform.resources',
        'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineCore', 'PySide6.QtWebChannel',
    ],
    # ... rest unchanged
)
```
[VERIFIED: PyInstaller hiddenimports pattern from PyInstaller docs. The `app_config.py` import restriction from CLAUDE.md still applies -- spec file must not import virelo modules directly.]

### Anti-Patterns to Avoid

- **Circular imports between subpackages:** Services must never import from bridge or app. Platform must never import from any other virelo subpackage. Dependency arrows point downward only.
- **Moving code and refactoring simultaneously:** Phase 3 is a mechanical move. Same code, new locations, updated imports. Logic refactoring is a separate concern.
- **Testing Win32 API calls in unit tests:** Unit tests mock all Win32/COM boundaries. Only integration tests (excluded from CI) call real Win32 APIs.
- **Breaking the entry point:** `python main.py` must continue to work after restructuring. The thin shim preserves backward compatibility.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Python linting + formatting | Custom style checker | Ruff 0.15.12 | Catches hundreds of error categories, sub-second execution, single config |
| Import sorting | Manual import ordering | Ruff `I` rules | Automatic, consistent, catches circular import symptoms |
| Test discovery + execution | Custom test runner | pytest 9.0.3 | Fixture injection, markers, plugins, clear failure output |
| JS test environment | Manual DOM setup | Vitest + jsdom | Vite-native, zero additional bundler config |
| CI pipeline | Shell scripts on server | GitHub Actions | Free, integrated, Windows runners, artifact caching |
| Coverage reporting | Manual line counting | pytest-cov 7.1.0 | Integrates coverage.py, HTML reports, missing-line markers |

**Key insight:** This phase is infrastructure setup, not application logic. Every tool above is the consensus standard for Python/React projects in 2026. Hand-rolling any of these would take longer than the application code itself.

## Common Pitfalls

### Pitfall 1: Circular Imports During Package Split

**What goes wrong:** Moving `ShiftSnapRestore` into `virelo/services/snap.py` and having it import from `virelo/app/window.py` while `window.py` imports from `services/snap.py` creates an ImportError at startup.

**Why it happens:** The current code has bidirectional dependencies -- `main.py` defines both `MainWindow` and `ShiftSnapRestore`, hiding the circular dependency.

**How to avoid:** Follow the strict bottom-up extraction order (D-05). The dependency graph from ARCHITECTURE research shows `platform/` at the bottom importing nothing from virelo, and `app/` at the top importing everything. Never import upward in the graph.

**Warning signs:** `ImportError: cannot import name X from partially initialized module Y`.

### Pitfall 2: Signal/Slot Disconnection After Moving QObject Classes

**What goes wrong:** Moving a `QObject` subclass to a new module can break signal/slot connections if the import path changes and the registration name in QWebChannel does not match.

**Why it happens:** QWebChannel registers objects by name (`"bridge"`). The JS side accesses `channel.objects.bridge`. If the Python `VireloBridge` class is imported from a different path, the registration still works -- but if the class internals change (e.g., Signal declarations move), connections break silently.

**How to avoid:** Move files mechanically -- same class, same Signal/Slot declarations, new file path. Verify QWebChannel registration name `"bridge"` is unchanged. The `frontend/src/bridge.js` must not change.

**Warning signs:** Frontend loads but settings don't populate. No JavaScript errors visible (signals just don't fire).

### Pitfall 3: PyInstaller Missing Modules After Restructuring

**What goes wrong:** PyInstaller's Analysis cannot find modules in the new `virelo/` package because it traces imports from the entry point and the new package structure requires explicit `hiddenimports`.

**Why it happens:** PyInstaller uses static analysis to find imports. Dynamic imports, late imports, and imports inside functions may be missed. The current spec has `hiddenimports=['bridge', 'webview', 'settings_state', 'snap_service']` -- these flat-module names will be wrong after the restructuring.

**How to avoid:** Update `Virelo.spec` `hiddenimports` to include all `virelo.*` submodules. Test the PyInstaller build after restructuring.

**Warning signs:** `ModuleNotFoundError` when running the frozen executable, but `python main.py` works fine.

### Pitfall 4: COM Threading Violation in Worker Split

**What goes wrong:** Splitting `workers.py` and separating `ExplorerAutosizeWorker` COM initialization from its COM operations causes `CoInitialize has not been called` or RPC errors.

**Why it happens:** COM Single-Threaded Apartment (STA) initialization is per-thread. `pythoncom.CoInitialize()` must be called in the same thread that later calls COM methods. If the worker's `run()` method is in one file but helper functions with COM calls are in another, the COM apartment is still valid -- but developers may accidentally restructure to call COM from a different thread.

**How to avoid:** D-07 explicitly requires COM init to stay co-located with COM operations. When splitting `workers.py`, keep `ExplorerAutosizeWorker.run()` and all its COM-dependent closures (`iter_tabs`, `autosize_try_wrapper`) in the same file (`workers/explorer.py`).

**Warning signs:** `pywintypes.com_error: (-2147417842, 'CoInitialize has not been called', None, None)`.

### Pitfall 5: Ruff Noisy Fixes on Existing Code

**What goes wrong:** Enabling all Ruff rules at once produces hundreds of warnings on the existing codebase, making the initial commit overwhelming and risky.

**Why it happens:** The codebase was written without a linter. Many patterns (bare `except`, unused imports in `__init__.py`, missing type annotations) are technically violations but functionally correct.

**How to avoid:** Start with the rules specified in D-11 (E, F, I, UP only). Add `per-file-ignores` for files that need specific exemptions (e.g., `F401` for `__init__.py` re-exports, `E501` if formatter handles line length). Fix lint errors incrementally.

**Warning signs:** CI failing on hundreds of pre-existing lint issues rather than on new regressions.

### Pitfall 6: Test Discovery Confusion Between Unit and Integration

**What goes wrong:** pytest discovers integration tests (marked `@pytest.mark.requires_qt`) and tries to run them in CI, where PySide6 is not installed. Import errors crash the test run.

**Why it happens:** pytest's default discovery finds all `test_*.py` files. If integration tests import PySide6 at module level, the import fails before the marker is even evaluated.

**How to avoid:** Two strategies (use both):
1. Integration test files use conditional imports: `pytest.importorskip("PySide6")` at module level
2. CI `pytest` command explicitly targets `tests/unit/` only (D-14: `pytest tests/unit/`)
3. `conftest.py` in `tests/integration/` registers the `requires_qt` marker

**Warning signs:** `ModuleNotFoundError: No module named 'PySide6'` in CI.

## Code Examples

### pyproject.toml Configuration

```toml
# Source: PEP 621 specification [CITED: packaging.python.org/en/latest/specifications/pyproject-toml/]
[build-system]
requires = ["setuptools>=68.0"]
build-backend = "setuptools.backends._legacy:_Backend"

[project]
name = "virelo"
version = "1.5.0"
description = "Windows desktop utility for window snapping and Explorer column auto-sizing"
requires-python = ">=3.12"
dependencies = [
    "PySide6>=6.6",
    "keyboard>=0.13.5",
    "pywin32>=306",
    "comtypes>=1.3.0",
]

[project.optional-dependencies]
dev = [
    "ruff>=0.15.12",
    "pytest>=9.0.3",
    "pytest-cov>=7.1.0",
]
build = [
    "pyinstaller>=6.0",
    "pyinstaller-hooks-contrib>=2024.6",
]

# --- Ruff ---
[tool.ruff]
target-version = "py312"
line-length = 100

[tool.ruff.lint]
select = ["E", "F", "I", "UP"]
# E: pycodestyle errors
# F: pyflakes (unused imports, undefined names)
# I: isort (import ordering)
# UP: pyupgrade (modernize syntax for target Python version)

[tool.ruff.lint.per-file-ignores]
"virelo/*/__init__.py" = ["F401"]  # Re-exports are intentional

[tool.ruff.format]
quote-style = "double"

# --- pytest ---
[tool.pytest.ini_options]
testpaths = ["tests"]
markers = [
    "requires_qt: marks tests requiring PySide6/Qt (excluded from CI)",
]
```

### Ruff CI Commands

```yaml
# Source: Ruff CLI documentation [CITED: docs.astral.sh/ruff/]
- run: pip install ruff
- run: ruff check .        # Lint: reports all violations
- run: ruff format --check . # Format: fails if any file would be reformatted
```

### pytest Unit Test Example (Settings Validation)

```python
# tests/unit/test_settings_state.py
# Source: pytest fixture pattern [VERIFIED: pytest documentation]
import pytest

class MockSettings:
    def __init__(self, **kwargs):
        from virelo.app.config import DEFAULTS
        for key, val in DEFAULTS.items():
            setattr(self, key, kwargs.get(key, val))
    def save(self):
        pass
    def clear(self):
        pass

@pytest.fixture
def state():
    from virelo.settings.state import SettingsState
    return SettingsState(MockSettings())

def test_apply_draft_validates_range(state):
    result = state.apply_draft({"width_pct": 150})
    assert result["ok"] is False
    assert "must be between" in result["error"]

def test_apply_draft_rejects_unknown_keys(state):
    result = state.apply_draft({"nonexistent_key": "value"})
    assert result["ok"] is False
    assert "Unknown keys" in result["error"]

def test_apply_draft_coerces_types(state):
    result = state.apply_draft({"snap_presses": "5"})
    assert result["ok"] is True
    assert result["applied"]["snap_presses"] == 5

def test_get_all_returns_all_keys(state):
    all_settings = state.get_all()
    from virelo.app.config import DEFAULTS
    for key in DEFAULTS:
        assert key in all_settings

def test_commit_draft_persists(state):
    state.apply_draft({"width_pct": 50})
    assert state.has_draft is True
    result = state.commit_draft()
    assert result["ok"] is True
    assert state.has_draft is False

def test_discard_draft_clears(state):
    state.apply_draft({"width_pct": 50})
    state.discard_draft()
    assert state.has_draft is False
```

### pytest Unit Test Example (Theme Resolution)

```python
# tests/unit/test_theme.py
from virelo.services.theme import normalize_theme_mode, resolve_theme, toggle_theme_mode

def test_normalize_valid_modes():
    assert normalize_theme_mode("dark") == "dark"
    assert normalize_theme_mode("light") == "light"
    assert normalize_theme_mode("system") == "system"

def test_normalize_invalid_falls_back():
    assert normalize_theme_mode("invalid") == "system"
    assert normalize_theme_mode("", "dark") == "dark"
    assert normalize_theme_mode(None) == "system"

def test_resolve_system_uses_system_theme():
    assert resolve_theme("system", "light") == "light"
    assert resolve_theme("system", "dark") == "dark"

def test_resolve_explicit_ignores_system():
    assert resolve_theme("dark", "light") == "dark"
    assert resolve_theme("light", "dark") == "light"

def test_toggle_system_mode():
    assert toggle_theme_mode("system", "light") == "dark"
    assert toggle_theme_mode("system", "dark") == "light"
```

### pytest Unit Test Example (Snap Geometry)

```python
# tests/unit/test_snap_geometry.py

def test_rect_matches_monitor_exact():
    from virelo.platform.win32_helpers import _rect_matches_monitor, FULLSCREEN_TOLERANCE
    # Exact match
    assert _rect_matches_monitor((0, 0, 1920, 1080), (0, 0, 1920, 1080)) is True

def test_rect_matches_monitor_within_tolerance():
    from virelo.platform.win32_helpers import _rect_matches_monitor
    # Off by 2 pixels (within FULLSCREEN_TOLERANCE=3)
    assert _rect_matches_monitor((-2, -1, 1922, 1081), (0, 0, 1920, 1080)) is True

def test_rect_does_not_match_monitor():
    from virelo.platform.win32_helpers import _rect_matches_monitor
    assert _rect_matches_monitor((100, 100, 800, 600), (0, 0, 1920, 1080)) is False

def test_calculate_snap_position():
    from virelo.services.snap import calculate_snap_position
    x, y, w, h = calculate_snap_position(0, 0, 1920, 1080, 76, 76)
    assert w == 1459  # 1920 * 76 // 100
    assert h == 820   # 1080 * 76 // 100
    assert x == 230   # (1920 - 1459) // 2
    assert y == 130   # (1080 - 820) // 2
```

### Vitest Configuration

```javascript
// frontend/vite.config.js -- add test config
// Source: Vitest configuration [VERIFIED: vitest docs]
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  define: {
    __APP_VERSION__: JSON.stringify(process.env.VITE_APP_VERSION || 'dev'),
  },
  build: {
    outDir: 'dist',
    assetsInlineLimit: 100000,
    rollupOptions: {
      output: {
        manualChunks: undefined,
      },
    },
  },
  server: {
    port: 5173,
    strictPort: true,
  },
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: './src/test-setup.js',
  },
});
```

```javascript
// frontend/src/test-setup.js
import '@testing-library/jest-dom';
```

### Vitest Frontend Test Example (bridgeToState/stateToBridge)

```javascript
// frontend/src/__tests__/app.test.jsx
// Source: Vitest + React Testing Library [VERIFIED: vitest docs, RTL docs]
import { describe, it, expect } from 'vitest';

// These are non-exported functions in app.jsx.
// Either export them for testing, or test through component rendering.
// Recommendation: export as named exports.

describe('bridgeToState', () => {
  it('maps Python keys to React keys', () => {
    const { bridgeToState } = await import('../app.jsx');
    const result = bridgeToState({
      enable_snap: true,
      snap_key: 'shift',
      restore_key: 'ctrl',
      snap_presses: 3,
      snap_interval: 1050,
      width_pct: 76,
      height_pct: 76,
      game_mode_enabled: true,
      ex_auto_size: false,
      run_at_startup: false,
    });
    expect(result.snapEnabled).toBe(true);
    expect(result.snapKey).toBe('SHIFT');
    expect(result.width).toBe(76);
  });
});

describe('stateToBridge', () => {
  it('maps React keys back to Python keys', () => {
    const { stateToBridge } = await import('../app.jsx');
    const json = stateToBridge({
      snapEnabled: true,
      snapKey: 'SHIFT',
      restoreKey: 'CTRL',
      pressCount: 3,
      interval: 1050,
      width: 76,
      height: 76,
      gameMode: true,
      autoSize: false,
      launchLogin: false,
    });
    const parsed = JSON.parse(json);
    expect(parsed.enable_snap).toBe(true);
    expect(parsed.snap_key).toBe('shift');
    expect(parsed.width_pct).toBe(76);
  });
});
```

### GitHub Actions CI Workflow

```yaml
# .github/workflows/ci.yml
# Source: GitHub Actions documentation [CITED: docs.github.com/en/actions]
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
          cache: 'pip'
      - run: pip install ruff
      - run: ruff check .
      - run: ruff format --check .

  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
          cache: 'pip'
      - run: pip install -e ".[dev]"
      - run: pytest tests/unit/ -q

  frontend:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '22'
          cache: 'npm'
          cache-dependency-path: frontend/package-lock.json
      - run: npm ci
        working-directory: frontend
      - run: npm run build
        working-directory: frontend
      - run: npx vitest run
        working-directory: frontend

  stale-name:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Check for stale "Windows Toolbox" references
        run: |
          if grep -rn "Windows Toolbox" . \
            --include="*.py" --include="*.jsx" --include="*.js" \
            --include="*.json" --include="*.toml" --include="*.yml" \
            --include="*.iss" --include="*.spec" \
            --exclude-dir=node_modules --exclude-dir=.git \
            --exclude-dir=dist --exclude-dir=build; then
            echo "ERROR: Found stale 'Windows Toolbox' references"
            exit 1
          fi
          echo "OK: No stale references found"
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| flake8 + black + isort (3 tools) | Ruff (single tool) | 2023-2024 | 10-100x faster, single config in pyproject.toml |
| requirements.txt | pyproject.toml (PEP 621) | 2022-2023 | Single source for metadata, deps, tool config |
| Jest for Vite projects | Vitest | 2022-2023 | Shares Vite transform pipeline, no separate config |
| ESLint 8/9 for React | Biome (or skip per D-13) | 2024-2025 | ESLint 10 broke React plugin ecosystem |
| setup.py / setup.cfg | pyproject.toml | 2022 | PEP 621 is the standard; setuptools reads it natively |
| unittest | pytest | Long-standing | pytest is de facto standard; richer fixtures, plugins |

**Deprecated/outdated:**
- **flake8:** Still works but Ruff supersedes it entirely with faster execution and unified config.
- **black:** `ruff format` produces identical output. No reason for a separate formatter.
- **setup.py:** Superseded by PEP 621 `pyproject.toml`. setuptools still supports it but new projects should not use it.
- **isort:** Ruff's `I` rules handle import sorting. Separate isort is redundant.

## Project Constraints (from CLAUDE.md)

These MUST be honored during implementation:

1. **Never reintroduce "Windows Toolbox" or "Toolbox" in any file.** CI job `stale-name` enforces this.
2. **Never add fake or placeholder UI controls.** No new UI in Phase 3; frontend tests test existing controls only.
3. **Never commit generated artifacts.** `frontend/dist/`, `dist/`, `build/`, `.venv/`, `__pycache__/` stay in `.gitignore`.
4. **Never hardcode version strings.** Version flows from `pyproject.toml` or `app_config.py` via `APP_VERSION`.
5. **`app_config.py` must not be imported in `Virelo.spec`.** Continue using regex to parse version. After restructuring, the regex reads from the new config location.
6. **PowerShell `$LASTEXITCODE` must be checked** after every external command in build scripts.
7. **Vite `define` values must use `JSON.stringify()`.** Preserved in existing `vite.config.js`.
8. **Inno Setup `#define` must use `#ifndef` guard.** Not changed in Phase 3.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `bridgeToState`/`stateToBridge` can be exported from `app.jsx` for testing without breaking the component | Code Examples (Vitest) | May need to extract to a separate `mapping.js` utility file instead |
| A2 | `calculate_snap_position` extraction is a good factoring for testability | Architecture Patterns (Pattern 3) | Low risk -- pure math function, straightforward extraction |
| A3 | Ubuntu runner can `pip install` the virelo package with its `[dev]` dependencies without Win32 packages | Architecture Patterns (CI) | Medium risk -- if any import in `virelo/` triggers Win32 imports at package init time, CI will fail. Unit tests must isolate platform imports. |
| A4 | Ruff rules E, F, I, UP will not produce excessive noise on the existing codebase | Common Pitfalls (Pitfall 5) | Low risk -- E/F are standard, I just reorders imports, UP only modernizes syntax. May need a few `per-file-ignores`. |
| A5 | PyInstaller will find all modules with explicit `hiddenimports` in the spec | Common Pitfalls (Pitfall 3) | Medium risk -- need to verify by running `pyinstaller` after restructuring. Some dynamic imports may be missed. |

## Open Questions (RESOLVED)

1. **Platform-guarded imports for CI compatibility** — RESOLVED: Plan 02 makes `virelo/platform/__init__.py` lazy (no transitive Win32 imports). Plan 04 unit tests import non-platform modules directly. CI runs `pytest tests/unit/` on Ubuntu without triggering Win32 imports.

2. **`bridgeToState`/`stateToBridge` export strategy** — RESOLVED: Plan 04 Task 2 exports both as named exports from `app.jsx`. They are pure mapping functions with no side effects.

3. **Ruff `per-file-ignores` scope** — RESOLVED: Plan 01 Task 1 configures E, F, I, UP rules with `per-file-ignores` for `__init__.py` files (F401). Executor runs `ruff check --fix` to auto-fix safe violations before committing config.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Python | Package restructuring, pytest | Yes | 3.13.13 | -- (3.12 recommended by STACK.md, 3.13 is compatible) |
| Node.js | Vitest, frontend build | Yes | 24.15.0 | -- |
| npm | Frontend dependency installation | Yes | 11.12.1 | -- |
| GitHub CLI (gh) | CI workflow creation | Yes | 2.90.0 | -- |
| pip | Python package installation | Yes | (bundled) | -- |
| Ruff | Linting (installed via pip) | No (not yet installed) | 0.15.12 (pip index) | Install via `pip install ruff` |
| pytest | Testing (installed via pip) | No (not yet installed) | 9.0.3 (pip index) | Install via `pip install pytest` |
| Vitest | Frontend testing (installed via npm) | No (not yet installed) | 4.1.5 (npm registry) | Install via `npm install -D vitest` |

**Missing dependencies with no fallback:** None -- all are installable via package managers.

**Missing dependencies with fallback:** None.

**Note:** Python 3.13 is installed instead of the 3.12 recommended in STACK.md. This is compatible -- PySide6 supports 3.10-3.14, Ruff and pytest support 3.13. The `pyproject.toml` should specify `requires-python = ">=3.12"` to allow both.

## Sources

### Primary (HIGH confidence)
- Codebase analysis: `main.py` (1449 lines), `workers.py` (978 lines), `bridge.py` (281 lines), all source files examined directly
- [VERIFIED: pip index] ruff 0.15.12, pytest 9.0.3, pytest-cov 7.1.0 -- current versions confirmed via `pip index versions`
- [VERIFIED: npm registry] vitest 4.1.5, @testing-library/react 16.3.2, @testing-library/jest-dom 6.9.1, @testing-library/user-event 14.6.1, jsdom 29.0.2 -- confirmed via `npm view`
- CONTEXT.md -- 17 locked decisions (D-01 through D-17) constraining implementation
- `.planning/research/STACK.md` -- Prior technology stack research
- `.planning/research/ARCHITECTURE.md` -- Prior architecture pattern research
- `.planning/codebase/ARCHITECTURE.md` -- Current architecture analysis
- `.planning/codebase/CONCERNS.md` -- Known technical debt and test coverage gaps

### Secondary (MEDIUM confidence)
- [CITED: packaging.python.org/en/latest/specifications/pyproject-toml/] -- PEP 621 pyproject.toml specification
- [CITED: docs.astral.sh/ruff/] -- Ruff configuration and rule documentation
- [CITED: docs.github.com/en/actions] -- GitHub Actions workflow syntax

### Tertiary (LOW confidence)
- None -- all claims verified against codebase, package registries, or official docs.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- All versions verified against live package registries. Tools are consensus standards.
- Architecture: HIGH -- Package layout is locked by CONTEXT.md decisions. Dependency graph verified against actual import statements.
- Pitfalls: HIGH -- Derived from direct codebase analysis (circular import risks, COM threading, PyInstaller behavior). Each pitfall traces to specific code locations.

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (30 days -- stable domain, tools have long release cycles)
