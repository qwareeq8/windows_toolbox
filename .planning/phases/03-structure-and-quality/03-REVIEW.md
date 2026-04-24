---
phase: 03-structure-and-quality
reviewed: 2026-04-24T21:30:00Z
depth: standard
files_reviewed: 36
files_reviewed_list:
  - virelo/__init__.py
  - virelo/app/__init__.py
  - virelo/app/__main__.py
  - virelo/app/config.py
  - virelo/app/webview.py
  - virelo/app/window.py
  - virelo/bridge/__init__.py
  - virelo/bridge/bridge.py
  - virelo/bridge/capture_guard.py
  - virelo/services/__init__.py
  - virelo/services/snap.py
  - virelo/services/explorer_columns.py
  - virelo/workers/__init__.py
  - virelo/workers/key_capture.py
  - virelo/workers/explorer.py
  - virelo/platform/__init__.py
  - virelo/platform/paths.py
  - virelo/platform/resources.py
  - virelo/platform/startup.py
  - virelo/platform/theme.py
  - virelo/platform/win32_helpers.py
  - virelo/settings/__init__.py
  - virelo/settings/persistence.py
  - virelo/settings/state.py
  - main.py
  - Virelo.spec
  - pyproject.toml
  - .github/workflows/ci.yml
  - tests/conftest.py
  - tests/integration/conftest.py
  - tests/unit/test_app_config.py
  - tests/unit/test_bridge_payload.py
  - tests/unit/test_capture_guard.py
  - tests/unit/test_paths.py
  - tests/unit/test_settings_state.py
  - tests/unit/test_snap_geometry.py
  - tests/unit/test_theme.py
  - frontend/src/__tests__/app.test.jsx
  - frontend/src/__tests__/panels.test.jsx
  - frontend/src/__tests__/primitives.test.jsx
  - frontend/src/test-setup.js
  - frontend/vite.config.js
  - frontend/package.json
findings:
  critical: 0
  warning: 4
  info: 3
  total: 7
status: issues_found
---

# Phase 03: Code Review Report

**Reviewed:** 2026-04-24T21:30:00Z
**Depth:** standard
**Files Reviewed:** 43
**Status:** issues_found

## Summary

Phase 03 restructured the entire Python backend from flat root-level modules into a `virelo/` package with six subpackages. The migration is architecturally sound: dependency arrows point in the right direction (platform -> settings -> services -> bridge -> app), `__init__.py` re-exports are consistent, the triple-dirname fix in `resources.py` is correct, and no circular imports were found.

The test suite is well-designed — native stubs in `conftest.py` are precise, the MockSettings fixture is clean, and the 50 Python / 20 JS tests cover meaningful behavior rather than just pass-through. CI pipeline jobs are correctly structured.

Four warnings are raised: a default-value mismatch between Python and the frontend (`ex_auto_size` / `autoSize`), a mismatched initial UI state in `app.jsx`, a global logger `setLevel` side-effect that will permanently escalate logging when Explorer autosize starts, and a CI pip cache that will never hit because there is no requirements file or lockfile on the pip path. Three info items are noted.

## Warnings

### WR-01: `ex_auto_size` default is `False` in Python but frontend defaults `autoSize` to `true`

**File:** `frontend/src/app.jsx:158` and `virelo/app/config.py:25`

**Issue:** `DEFAULTS["ex_auto_size"]` is `False` in Python (correct: Explorer autosize is off by default). However, `bridgeToState` maps `settings.ex_auto_size ?? true`, using `true` as the JavaScript nullish-coalescing fallback. This means that if the bridge ever sends `null` or `undefined` for `ex_auto_size` — or if `bridgeToState` is called before the first `get_settings` response arrives — the UI will show Explorer autosize as enabled when it is actually disabled. Additionally, the `VireloApp` component initial state also hardcodes `autoSize: true` (line 187), compounding the mismatch.

In practice the initial state is overwritten by the first `get_settings` call, so users won't see a persistent wrong value. But the window will briefly render with `autoSize: true` before the bridge responds, and any test calling `bridgeToState({})` will get `autoSize: true` — which the existing test at line 46 asserts as correct despite contradicting the Python default.

**Fix:**
```js
// frontend/src/app.jsx
// bridgeToState: change default from true to false
autoSize: settings.ex_auto_size ?? false,

// VireloApp initial state: match Python DEFAULTS
const [state, setState] = React.useState({
  snapEnabled: true, snapKey: 'SHIFT', restoreKey: 'CTRL',
  pressCount: 3, interval: 1050, width: 76, height: 76,
  gameMode: true, autoSize: false,    // matches DEFAULTS["ex_auto_size"] = False
  launchLogin: false,
});
```
Update `app.test.jsx` line 46 to `expect(result.autoSize).toBe(false)` to match.

---

### WR-02: `VireloApp` initial state sets `launchLogin: true` but `DEFAULTS["run_at_startup"]` is `False`

**File:** `frontend/src/app.jsx:188`

**Issue:** The initial React state has `launchLogin: true`, but the Python default for `run_at_startup` is `False`. The `bridgeToState` nullish fallback for this field (`?? false`) is correct. The discrepancy is only in the static `useState` initializer, meaning the window flickers with "run at startup" appearing checked until the bridge responds. Unlike WR-01 this is a single-field issue with no test assertion locking it in.

**Fix:**
```js
const [state, setState] = React.useState({
  ...
  launchLogin: false,  // matches DEFAULTS["run_at_startup"] = False
});
```

---

### WR-03: `LOG.setLevel(logging.DEBUG)` in `_update_explorer_autosize_thread` mutates global logger level permanently

**File:** `virelo/app/window.py:445`

**Issue:** When Explorer autosize is enabled, `_update_explorer_autosize_thread` calls `LOG.setLevel(logging.DEBUG)` on the module-level `LOG` logger. This permanently escalates the `"Virelo"` logger to DEBUG for the rest of the process lifetime — every subsequent log call from any module sharing this logger will produce DEBUG-level output, flooding the log file even when autosize is later disabled. The logger was already initialized at DEBUG level in `__main__.py` line 21 (with the file handler at DEBUG), so this call is redundant and its only visible effect is disabling any runtime log-level adjustment the operator may have configured.

**Fix:** Remove the `setLevel` call from `_update_explorer_autosize_thread`. The file handler already captures DEBUG-level entries since `__main__.py` sets `logger.setLevel(logging.DEBUG)` at startup:
```python
# Remove these two lines from _update_explorer_autosize_thread:
# LOG.setLevel(logging.DEBUG)
# LOG.info("Explorer autosize: enabling DEBUG logging for troubleshooting")
```

---

### WR-04: CI pip cache will never hit — `setup.cfg` or `requirements*.txt` absent, and `cache: 'pip'` requires them

**File:** `.github/workflows/ci.yml:17,29`

**Issue:** The `lint` and `test` jobs use `actions/setup-python@v5` with `cache: 'pip'`. The `pip` cache strategy in this action requires a dependency file (`requirements*.txt`, `setup.cfg`, or `pyproject.toml` at a recognized path) to compute the cache key. With `pyproject.toml` present this will likely work for the `test` job since `pip install -e ".[dev]"` reads `pyproject.toml`. However, the `lint` job only runs `pip install ruff` with no dependency file involvement — pip caching here requires a `requirements-*.txt` pattern or explicit `cache-dependency-path`. The cache key will be stale or uncacheable, causing a cache miss every run.

This is a performance issue (extra network latency), not a correctness failure — CI will still pass. However, it degrades CI speed on every run.

**Fix:**
```yaml
# In the lint job, add cache-dependency-path or pin ruff via a requirements file.
# Simplest: add cache-dependency-path for both jobs:
- uses: actions/setup-python@v5
  with:
    python-version: '3.12'
    cache: 'pip'
    cache-dependency-path: 'pyproject.toml'   # add this line to both jobs
```

## Info

### IN-01: `virelo/app/__init__.py` imports `main` from `__main__`, creating an unusual import path

**File:** `virelo/app/__init__.py:1`

**Issue:** `from virelo.app.__main__ import main` imports a function from a `__main__` module, which is an unusual pattern. Python's `__main__` module name has special meaning — it is the entry point when run as `python -m virelo.app`. While this works correctly, it creates a subtle confusion: `from virelo.app import main` actually imports from `virelo/app/__main__.py`, not from a conventional module. If someone later adds `if __name__ == "__main__": main()` to `__main__.py` and also uses `from virelo.app import main`, the double-import is harmless but the intent is obscured.

**Fix (optional):** Move `main()` to `virelo/app/_entry.py` or `virelo/app/main.py` and keep `__main__.py` as the thin `if __name__ == "__main__": from virelo.app import main; main()` shim. This separates the runnable entry point from the callable function. Not a blocking issue given the current working implementation.

---

### IN-02: `test_capture_guard.py` thread-safety test is deterministic only under the Barrier assumption

**File:** `tests/unit/test_capture_guard.py:29-45`

**Issue:** `test_concurrent_try_start` uses a `threading.Barrier(10)` to synchronize 10 threads before they call `guard.try_start()`. The test asserts `results.count(True) == 1`. This is correct given the `CaptureGuard` implementation, but the test relies on all 10 threads reaching the barrier — if any thread fails to start (e.g., resource exhaustion), `Barrier.wait()` would raise `BrokenBarrierError` and the test would fail with an unclear error rather than a meaningful assertion failure. This is a minor reliability edge case in very constrained CI environments.

**Fix (optional):** Wrap the test in a `try/except BrokenBarrierError` or reduce thread count. Not urgent.

---

### IN-03: `pyproject.toml` version `1.5.0` is hardcoded separately from `virelo/app/config.py`

**File:** `pyproject.toml:7`

**Issue:** `pyproject.toml` contains `version = "1.5.0"` as a static string. `virelo/app/config.py` also contains `APP_VERSION = "1.5.0"`. CLAUDE.md Forbidden Change #4 says "Never hardcode version strings" — use `APP_VERSION` from `virelo/app/config.py`. The `pyproject.toml` version cannot easily read from the Python module (it's metadata, not code), but it creates a second place to update at release time. If someone bumps `APP_VERSION` in `config.py` but forgets `pyproject.toml`, the package metadata diverges from the application version.

**Fix:** Use `setuptools-scm` or `hatch-vcs` to derive `pyproject.toml` version from a single source, or add a release checklist note that both files must be updated together. Alternatively, use `dynamic = ["version"]` with a `tool.setuptools.dynamic` that reads `virelo/app/config.py` via a regex (same technique used in `Virelo.spec`).

---

_Reviewed: 2026-04-24T21:30:00Z_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
