# Testing Patterns

**Analysis Date:** 2026-04-24

## Test Framework

**Runner:**
- None configured. No test framework is installed or set up.
- No `pytest.ini`, `conftest.py`, `tox.ini`, `jest.config.*`, or `vitest.config.*` detected.
- No test dependencies in `requirements.txt` or `frontend/package.json`.

**Assertion Library:**
- Not applicable -- no test infrastructure exists.

**Run Commands:**
```bash
# No test commands available
# Python: no pytest/unittest/nose configured
# JavaScript: no jest/vitest/mocha configured
```

## Test File Organization

**Location:**
- No test files exist anywhere in the project.
- No `tests/` directory, no `__tests__/` directory, no `*.test.*` or `*.spec.*` files.
- No `test_*.py` or `*_test.py` files.

**One defensive coding marker found:**
- `workers.py` line 632: `except Exception:  # pragma: no cover - PySide6 unavailable in some test envs`
- This comment suggests tests may have existed at some point or were planned, but none are present.

## Test Structure

**No tests exist.** The sections below describe what patterns should be adopted if tests are added, based on the codebase architecture.

## Recommended Test Approach

**Testable units by design:**

Several modules use dependency injection that makes them testable without mocking frameworks:

1. **`theme.py`** -- Pure functions, fully testable:
   ```python
   # theme.py functions accept optional callable for registry access
   def get_windows_theme(read_registry=None):
       # If read_registry is None, reads from Windows registry
       # If provided, calls read_registry() instead
   ```
   Test: `assert get_windows_theme(read_registry=lambda: 1) == "light"`

2. **`startup_shortcut.py`** -- Pure functions with injected `exists`:
   ```python
   def startup_shortcut_spec(executable, argv0, frozen, exists=os.path.exists):
   ```
   Test: `assert startup_shortcut_spec("python.exe", "main.py", False, exists=lambda p: False) == ("python.exe", '"main.py"')`

3. **`settings_state.py`** -- Validation logic with clear inputs/outputs:
   ```python
   # apply_partial returns {"ok": True/False, ...}
   result = state.apply_partial({"snap_presses": 5})
   assert result["ok"] is True
   ```

4. **`app_config.py`** -- Simple normalize functions:
   ```python
   assert normalize_snap_presses(0) == 1  # clamped to min 1
   assert normalize_snap_presses("abc") == 3  # falls back to default
   ```

5. **`capture_guard.py`** -- Thread-safe guard, testable:
   ```python
   guard = CaptureGuard()
   assert guard.try_start() is True
   assert guard.try_start() is False  # already active
   guard.finish()
   assert guard.try_start() is True
   ```

6. **`workers.py` `KeyCaptureSession`** -- Accepts `keyboard` module as dependency:
   ```python
   session = KeyCaptureSession(mock_keyboard, cancel_key="esc", timeout_s=0.1)
   ```

7. **`workers.py` `ExplorerAutosizeEngine`** -- Accepts all callables as constructor arguments:
   ```python
   engine = ExplorerAutosizeEngine(
       iter_tabs=lambda: [],
       autosize_try=lambda hwnd, path: (True, "mock", False),
       autosize_full=lambda hwnd, path: (True, "mock", False),
       is_window_interactive=lambda hwnd: True,
   )
   delay = engine.step(time.time())
   ```

**Hard-to-test units:**

- `main.py` `MainWindow` -- Tightly coupled to Qt event loop, Win32 APIs, system tray, WebView
- `main.py` `ShiftSnapRestore` -- Depends on `keyboard` hooks, `win32gui.EnumWindows`, live window handles
- `explorer_columns.py` -- COM interface calls to Shell.Application, requires running Explorer windows
- `webview.py` -- Requires QWebEngine runtime
- `bridge.py` -- Depends on `SettingsState`, `SnapService`, and `MainWindow` at runtime

## Mocking

**Framework:** Not applicable (no test framework installed)

**Recommended approach if adding tests:**
- Use `unittest.mock` (stdlib) for Python
- The codebase already uses dependency injection in several modules, reducing the need for heavy mocking
- For COM-dependent code (`explorer_columns.py`, `workers.py`), mock the entire Shell.Application interface
- For Qt-dependent code, consider `pytest-qt` with `QApplication` fixture

**What to mock:**
- Win32 API calls: `win32gui`, `win32api`, `ctypes.windll`
- COM objects: `comtypes.client.CreateObject`, `Shell.Application`
- `keyboard` module for capture tests
- File system operations for startup shortcut tests
- QSettings for settings persistence tests

**What NOT to mock:**
- Pure functions in `theme.py`, `app_config.py`, `startup_shortcut.py` -- test directly
- Validation logic in `settings_state.py` -- instantiate with a mock `Settings` object
- `ExplorerAutosizeEngine` -- use injected callables, no mocking needed
- `CaptureGuard` -- simple threading primitive, test directly

## Fixtures and Factories

**Test Data:**
- Default settings values are defined in `app_config.py` `DEFAULTS` dict -- use as test fixture:
  ```python
  DEFAULTS = {
      "snap_key": "shift",
      "restore_key": "ctrl",
      "enable_snap": True,
      "snap_presses": 3,
      "snap_interval": 1050,
      "width_pct": 76,
      "height_pct": 76,
      "ex_auto_size": False,
      "game_mode_enabled": True,
      "run_at_startup": False,
      "theme": "system",
  }
  ```
- Mock bridge settings in `frontend/src/bridge.js` `MOCK_SETTINGS` -- use as JS test fixture

**Location:**
- Not applicable -- no test fixtures exist

## Coverage

**Requirements:** None enforced. No coverage tool configured.

**Recommended setup:**
```bash
# Python
pip install pytest pytest-cov
pytest --cov=. --cov-report=html

# JavaScript
npm install -D vitest @testing-library/react
npx vitest --coverage
```

## Test Types

**Unit Tests:**
- Not present. Priority targets for unit testing:
  - `theme.py` -- 4 pure functions, zero dependencies
  - `app_config.py` -- `normalize_snap_presses()` validation
  - `settings_state.py` -- `apply_partial()` validation and coercion
  - `capture_guard.py` -- thread-safety of `try_start()`/`finish()`
  - `startup_shortcut.py` -- path selection logic
  - `workers.py` `ExplorerAutosizeEngine` -- state machine logic (debounce, dedup, circuit breaker)
  - `workers.py` `_canonicalize_path()` -- path normalization edge cases

**Integration Tests:**
- Not present. Would require:
  - Qt event loop (`pytest-qt`) for bridge/webview tests
  - COM environment for Explorer column tests
  - System tray availability for MainWindow tests

**E2E Tests:**
- Not present. The application runs as a desktop app with WebView frontend.
- Potential approach: Selenium/Playwright driving the embedded Chromium via remote debugging, or Qt Test framework

## Common Patterns

**Async Testing (if added):**
```python
# Workers use threading.Event for synchronization
# Test pattern for KeyCaptureSession:
session = KeyCaptureSession(mock_keyboard, timeout_s=0.5)
result, reason = session.run()
assert reason == "timeout"
```

**Error Testing (if added):**
```python
# Bridge methods return JSON with ok/error fields
# Test pattern for validation:
result = json.loads(bridge.save_settings('{"snap_presses": 999}'))
assert result["ok"] is False
assert "must be between" in result["error"]

# Test pattern for invalid JSON:
result = json.loads(bridge.save_settings('not json'))
assert result["ok"] is False
assert "Invalid JSON" in result["error"]
```

## CI/CD

**Pipeline:** None configured.
- No `.github/workflows/` directory
- No `.gitlab-ci.yml`
- No `Jenkinsfile`
- Build is manual via PyInstaller (`Windows Toolbox.spec`) and Inno Setup (`installer/virelo.iss`)
- Frontend build: `npm run build` (Vite) in `frontend/`

**Recommended CI setup:**
```yaml
# Minimum viable CI:
# 1. Lint Python (ruff) and JS (eslint)
# 2. Run unit tests for pure-function modules
# 3. Build frontend (vite build)
# 4. Build executable (pyinstaller)
```

---

*Testing analysis: 2026-04-24*
