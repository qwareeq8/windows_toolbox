# Codebase Concerns

**Analysis Date:** 2026-04-24

## Tech Debt

**Monolithic `main.py` (1451 lines) - God File:**
- Issue: `main.py` contains Win32 API constants, helper functions (~30 utility functions for window management), the `ShiftSnapRestore` class (290 lines), the entire `MainWindow` class (540+ lines), startup shortcut management, and the entry point. All in a single file.
- Files: `main.py`
- Impact: Difficult to navigate, test in isolation, or modify safely. Any change to window helpers risks regretting snap logic or vice versa.
- Fix approach: Extract into modules:
  - `win32_helpers.py` for `_get_rect`, `_area`, `_ancestor_classes`, `_find_descendant_by_class`, `_collect_descendants_by_class`, `_is_window_interactive`, `_looks_like_preview`, `_find_best_folder_listview`, `get_monitor_rect`, `_get_window_dwm_rect`, `_is_window_fullscreen`, `_looks_like_game_window`, `_should_skip_snap_for_game`, `_exit_fullscreen`, `_as_hwnd`, `_get_children`, `_class_name`
  - `snap_restore.py` for `ShiftSnapRestore`
  - `startup.py` for `get_startup_shortcut_path`, `create_startup_shortcut`, `remove_startup_shortcut`
  - Keep `MainWindow` and entry point in `main.py`

**Monolithic `workers.py` (978 lines) - Mixed Concerns:**
- Issue: Contains `KeyCaptureSession` (pure threading logic), `_canonicalize_path` (utility), `ExplorerAutosizeEngine` (stateful engine, ~400 lines), path resolution helpers, PySide6 worker classes (`KeyCaptureWorker`, `ExplorerAutosizeWorker` with 250-line `run` method containing deeply nested closures).
- Files: `workers.py`
- Impact: The `ExplorerAutosizeWorker.run()` method defines `iter_tabs`, `autosize_try_wrapper`, `autosize_full_wrapper` as closures that capture `shell_app` and `self._stop`. This makes testing and debugging extremely difficult.
- Fix approach: Extract `ExplorerAutosizeEngine` into its own module. Promote `iter_tabs` from a closure to a proper method or class. Extract `KeyCaptureSession` to a standalone module.

**Monolithic `explorer_columns.py` (861 lines) - Low-Level COM + High-Level Logic:**
- Issue: Mixes low-level COM interface definitions (ctypes structures, IColumnManager, IServiceProvider), window enumeration, tab management, path canonicalization, and autosize orchestration in one file.
- Files: `explorer_columns.py`
- Impact: Hard to understand or modify the COM layer independently from the tab-management logic.
- Fix approach: Split into `com_interfaces.py` (structures, interface definitions), `shell_windows.py` (enumeration, tab finding), and `column_autosize.py` (autosize logic).

**Duplicate Path Canonicalization - Three Copies:**
- Issue: Path canonicalization logic exists in three separate implementations with slightly different behavior:
  1. `explorer_columns.py:850` - `canonicalize_path()` (public export)
  2. `workers.py:143` - `_canonicalize_path()` (private, used by engine)
  3. `explorer_columns.py:402` - `normalize_path()` (local function inside `find_explorer_tab_by_path`)
- Files: `explorer_columns.py`, `workers.py`
- Impact: Subtle path comparison bugs if implementations diverge. Currently they handle `file:///` URLs, backslash normalization, and case folding slightly differently.
- Fix approach: Consolidate into a single `canonicalize_path()` in a shared utility module and import everywhere.

**Duplicate `resource_path` Functions:**
- Issue: Two identical implementations of `resource_path` exist:
  1. `main.py:198` - `resource_path()`
  2. `webview.py:33` - `_resource_path()` (private copy with comment: "inline copy to avoid circular import")
- Files: `main.py`, `webview.py`
- Impact: If one is updated and the other is not, PyInstaller resource resolution breaks in one context. The circular import concern is valid but solvable.
- Fix approach: Move `resource_path` to a utility module (e.g., `app_config.py` or a new `paths.py`) that both `main.py` and `webview.py` import.

**Duplicate Tab State Dataclasses:**
- Issue: Two similar but distinct tab state dataclasses:
  1. `explorer_columns.py:185` - `TabState` (7 fields)
  2. `workers.py:49` - `TabAutosizeState` (15 fields, superset)
- Files: `explorer_columns.py`, `workers.py`
- Impact: `TabState` in `explorer_columns.py` appears unused by the active autosize engine (which uses `TabAutosizeState` from `workers.py`). Dead code that confuses maintainers.
- Fix approach: Remove `TabState` from `explorer_columns.py` if unused, or consolidate into a single dataclass.

**Duplicate Explorer Path Resolution:**
- Issue: `_resolve_explorer_path()` in `workers.py:600` duplicates path resolution logic that `explorer_columns.py` also handles via `iter_shell_windows()` and `ShellWindow.location_url`.
- Files: `workers.py:600-627`
- Impact: Two codepaths for resolving Explorer paths from COM objects. Maintenance burden.
- Fix approach: Use `explorer_columns.py`'s path resolution consistently.

**Identical Quick/Full Autosize Functions:**
- Issue: `_autosize_explorer_columns_quick()` and `_autosize_explorer_columns_full()` at `main.py:531-554` have identical implementations. The docstring for `_full` says "currently identical to quick."
- Files: `main.py:531-554`
- Impact: Two functions that do the same thing, creating confusion about when each should be used. The engine selects between them based on retry count but both call `autosize_explorer_columns(..., allow_keyboard_fallback=False, ...)`.
- Fix approach: Remove the distinction or implement actual fallback behavior in `_full`.

**Empty `_update_snap_enabled_state` Method:**
- Issue: `MainWindow._update_snap_enabled_state()` at line 1199 is `pass` with a comment "React reads snap_enabled from settings via bridge." This is called from multiple locations (`bridge.py:96`, `bridge.py:208`, `main.py:1171`, `main.py:1280`).
- Files: `main.py:1199-1200`
- Impact: Dead code paths that obscure intent. Callers think they're updating state but nothing happens.
- Fix approach: Either remove all call sites or implement the method to emit a bridge signal if React actually needs a push notification.

## Security Considerations

**Admin Elevation Without User Confirmation:**
- Risk: Application unconditionally requests admin elevation on every launch via `ShellExecuteW("runas", ...)` at `main.py:1397-1411`. If the process is not already admin, it re-launches itself elevated without asking the user first.
- Files: `main.py:1396-1412`
- Current mitigation: Windows UAC prompt is shown (OS-level). Single-instance mutex prevents duplicate processes.
- Recommendations: Document why admin rights are needed (global keyboard hooks, COM access to Explorer). Consider running without admin by default and only requesting elevation for specific features that require it (e.g., keyboard hooks on elevated windows).

**Global Keyboard Hook:**
- Risk: The `keyboard` library installs a system-wide keyboard hook (`keyboard.on_press_key`, `keyboard.hook`) that captures all keystrokes across all applications. Combined with admin elevation, this has full access to every keystroke on the system.
- Files: `main.py:626-627` (ShiftSnapRestore constructor), `workers.py:100` (KeyCaptureSession)
- Current mitigation: The hook only processes the configured snap/restore keys and discards everything else. Key capture sessions have a 15-second timeout.
- Recommendations: Add clear documentation about the global hook. Consider using a Windows low-level keyboard hook directly (via `ctypes`) with stricter filtering instead of the `keyboard` library's global hook.

**Registry Read (Low Risk):**
- Risk: Theme detection reads from `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize`. Read-only, user-scoped.
- Files: `theme.py:26-32`
- Current mitigation: Wrapped in try/except, falls back to "dark" on any error.
- Recommendations: No action needed. Read-only access to user's own registry is safe.

**No `.gitignore` File:**
- Risk: No `.gitignore` file exists in the project root. Any new sensitive file (`.env`, credentials, build artifacts, `__pycache__`) could be accidentally committed.
- Files: Project root
- Current mitigation: None detected.
- Recommendations: Create a `.gitignore` covering `__pycache__/`, `*.pyc`, `*.pyo`, `.env`, `dist/`, `build/`, `*.spec`, `frontend/node_modules/`, `frontend/dist/`, `*.egg-info/`, `venv/`.

## Performance Bottlenecks

**COM Object Creation Per Poll in `explorer_columns.py`:**
- Problem: `iter_shell_windows()` at `explorer_columns.py:204` calls `comtypes.client.CreateObject("Shell.Application")` on every invocation. This is used by `find_explorer_tab_by_path`, `find_active_tab_for_hwnd`, etc.
- Files: `explorer_columns.py:204-242`
- Cause: No caching of the Shell.Application COM object in the `explorer_columns.py` module functions. The worker thread in `workers.py` caches its own Shell.Application at line 741, but the `explorer_columns.py` functions called during autosize do not share that cache.
- Improvement path: Accept a Shell.Application instance as a parameter, or provide a context manager that caches it for the duration of an autosize operation.

**BFS Window Traversal With Deep Nesting:**
- Problem: `_find_descendant_by_class()` and `_collect_descendants_by_class()` in `main.py:424-472` use BFS with `queue.pop(0)` (O(n) list pop from front) to depth 12-14 levels.
- Files: `main.py:424-472`
- Cause: Using a list as a queue instead of `collections.deque`. Also, deep traversal of the window tree for every snap operation.
- Improvement path: Replace `queue.pop(0)` with `collections.deque.popleft()`. Consider caching the SHELLDLL_DefView handle per Explorer window.

**Debug Logging Left Enabled in Production:**
- Problem: At `main.py:1228`, when the Explorer autosize feature is enabled, `LOG.setLevel(logging.DEBUG)` is called, forcing ALL logging across the application to DEBUG level. This generates high-volume log output during normal operation.
- Files: `main.py:1228-1230`
- Cause: Troubleshooting code left in place. The logger is set to DEBUG globally, not just for the explorer subsystem.
- Improvement path: Remove the `LOG.setLevel(logging.DEBUG)` call. The file handler is already configured at DEBUG level (line 91); the logger root level is already DEBUG (line 70). This additional call is redundant and its comment suggests it was temporary troubleshooting.

**Polling Loop With 5ms Sleep:**
- Problem: The `ExplorerAutosizeWorker.run()` main loop at `workers.py:960-968` uses `time.sleep(0.005)` in a tight loop for interruptible sleep. When no Explorer windows are open, it still polls every ~150ms.
- Files: `workers.py:925-968`
- Cause: COM STA apartment threading requires message pumping, so the worker cannot use a simple blocking wait.
- Improvement path: Increase the idle poll interval when no Explorer windows are detected. Currently 150ms when no pending work; could safely be 500ms-1s.

## Fragile Areas

**Explorer Tab Detection Heuristics:**
- Files: `explorer_columns.py:291-348`, `workers.py:760-879`
- Why fragile: Windows 11 tab support relies on heuristics: scoring tabs by whether they have a URL and are in Details view, because there is no reliable COM API to determine the active tab. The focus-detection code at `explorer_columns.py:321-329` contains a dead code path (the `if focus_hwnd:` block does nothing with the result).
- Safe modification: Any changes must be tested on both Windows 10 (no tabs) and Windows 11 (with tabs). The `_get_focus_hwnd` function at `explorer_columns.py:254-267` uses raw struct manipulation with hardcoded byte offsets (`gui_info[8:16]`) for the `GUITHREADINFO` structure, which is architecture-sensitive.
- Test coverage: No automated tests exist for any Explorer interaction code.

**Bridge/MainWindow Tight Coupling:**
- Files: `bridge.py`, `main.py:905-1389`
- Why fragile: `VireloBridge` calls private methods on `MainWindow` directly (e.g., `self._main_window._start_key_capture()`, `self._main_window._update_snap_enabled_state()`, `self._main_window._apply_theme_mode()`). There are 7 direct calls to `_main_window._*` private methods in `bridge.py`.
- Safe modification: Changes to `MainWindow`'s private API require checking all bridge call sites. Adding new bridge slots requires modifying both `bridge.py` and `main.py`.
- Test coverage: No tests for bridge-to-MainWindow interaction.

**COM Lifecycle in Worker Thread:**
- Files: `workers.py:710-977`
- Why fragile: The `ExplorerAutosizeWorker.run()` method manages COM initialization, Shell.Application caching, message pumping, and cleanup in a complex try/finally block with multiple nested closures. The `iter_tabs` closure checks `self._stop.is_set()` before every COM call (12 separate checks) to avoid RPC failures during shutdown.
- Safe modification: Any changes to the worker loop must preserve the order of: COM init -> Shell.Application creation -> engine loop -> Shell.Application release -> COM uninit. Reordering causes hard crashes.
- Test coverage: No automated tests. The `# pragma: no cover` at `workers.py:632` explicitly acknowledges this.

**Keyboard Hook Lifecycle:**
- Files: `main.py:626-637`, `main.py:691-702`
- Why fragile: `ShiftSnapRestore` hooks keyboard events in the constructor and unhooks in `cleanup()`. If `update_binding()` is called rapidly, old hooks might not be cleanly removed (the `except Exception: pass` at lines 693-699 silently swallows unhook failures). The `keyboard` library's hook management is global state.
- Safe modification: Always call `cleanup()` before destruction. Consider adding a lock around hook registration/unregistration.
- Test coverage: No automated tests for hook lifecycle.

**`_get_focus_hwnd` Raw Struct Manipulation:**
- Files: `explorer_columns.py:254-267`
- Why fragile: Manually constructs a `GUITHREADINFO` structure using `ctypes.create_string_buffer(48)` with hardcoded size and byte offset `[8:16]` for the `hwndFocus` field. This assumes specific struct layout and pointer size (64-bit Windows).
- Safe modification: Replace with a proper ctypes Structure definition for `GUITHREADINFO` to avoid byte-offset fragility.
- Test coverage: None.

## Scaling Limits

**In-Memory Deduplication Cache (Unbounded Growth Potential):**
- Current capacity: `ExplorerAutosizeEngine._dedupe_cache` is cleaned periodically (5-minute TTL) and `autosized_paths` is bounded to 500 entries (LRU).
- Limit: The `tab_state` dictionary grows with the number of unique `(hwnd, path)` combinations. Dead entries are cleaned when tabs close, but rapid navigation could create many entries.
- Scaling path: Current limits are adequate for typical usage (tens of Explorer windows). No action needed unless users report memory issues.

**Single-Instance Mutex:**
- Current capacity: One instance of Virelo per Windows user session.
- Limit: The mutex at `main.py:1425` uses `Global\` prefix, meaning it is per-session, not per-user. On multi-user RDP hosts, only one user can run Virelo.
- Scaling path: Change to a per-user mutex if multi-user RDP support is needed.

## Dependencies at Risk

**`keyboard` Library (v0.13.5+):**
- Risk: The `keyboard` library requires root/admin on all platforms and installs a global keyboard hook. It has known issues with certain keyboard layouts and sometimes conflicts with other keyboard hook software. Last significant release was 2020.
- Impact: Core snap/restore functionality depends on it. Admin elevation requirement is partially driven by this library.
- Migration plan: Consider switching to a Windows-specific low-level keyboard hook via `ctypes` and `SetWindowsHookExW`, which would give more control and potentially remove the admin requirement for non-elevated windows.

**PySide6 (v6.6+) - Qt WebEngine Dependency:**
- Risk: PySide6 bundles Chromium via Qt WebEngine, adding ~200MB to the distribution. The `AA_EnableHighDpiScaling` and `AA_UseHighDpiPixmaps` attributes used at `main.py:1420-1421` are deprecated in Qt 6 (they are always enabled). This causes deprecation warnings.
- Impact: Build size and startup time. Deprecation warnings in logs.
- Migration plan: Remove the two `setAttribute` calls at `main.py:1420-1421`. They are no-ops in Qt 6 and generate warnings.

**`comtypes` Library (v1.3.0+):**
- Risk: `comtypes` generates Python wrappers in a temp directory (`gen_py` cache). This cache can become corrupted, which is why `main.py:141-162` includes `_ensure_dispatch()` with cache-rebuilding logic.
- Impact: If the gen_py cache corrupts, COM operations fail until the cache is rebuilt. This can happen after Windows updates.
- Migration plan: The existing workaround (`_ensure_dispatch`) handles this. No urgent action needed.

## Missing Critical Features

**No Test Suite:**
- Problem: Zero test files exist in the project. No `pytest.ini`, `conftest.py`, `tox.ini`, `pyproject.toml`, or any `test_*.py` files. The `# pragma: no cover` comment at `workers.py:632` confirms testing was considered but not implemented.
- Blocks: Safe refactoring, CI/CD, regression detection. Any of the structural improvements listed above (file splitting, deduplication) are risky without tests.

**No `.gitignore`:**
- Problem: No `.gitignore` file exists. Not a git repo currently, but when one is created, build artifacts (`__pycache__`, `dist/`, `build/`, `frontend/node_modules/`, `frontend/dist/`) would be tracked.
- Blocks: Clean version control.

**No CI/CD Pipeline:**
- Problem: No GitHub Actions, Azure Pipelines, or other CI configuration exists. No automated build, test, or release process.
- Blocks: Automated quality checks, reproducible builds, release automation.

**No Type Checking Configuration:**
- Problem: No `mypy.ini`, `pyrightconfig.json`, or `py.typed` marker. Type hints exist in some files (`main.py` uses `Optional`, `Tuple`, `Dict`, etc.) but are not enforced.
- Blocks: Catching type errors at development time.

## Test Coverage Gaps

**All Code - Zero Test Coverage:**
- What's not tested: The entire codebase. No test files exist.
- Files: All `.py` files
- Risk: Every area is at risk of undetected regressions. The most critical untested areas are:
  1. `ShiftSnapRestore` snap/restore logic (`main.py:609-894`) - window manipulation that could move/resize windows incorrectly
  2. `ExplorerAutosizeEngine` state machine (`workers.py:174-584`) - complex state transitions with circuit breakers, deduplication, and rate limiting
  3. `Settings` persistence (`settings.py`) - data loss if save/load is broken
  4. `SettingsState` validation (`settings_state.py`) - could allow invalid settings if bounds checking fails
  5. `VireloBridge` slot methods (`bridge.py`) - all frontend-to-backend communication
  6. Path canonicalization (3 implementations) - could cause autosize to fire repeatedly or miss navigations
- Priority: High. Adding tests for `SettingsState`, `ExplorerAutosizeEngine.step()`, path canonicalization, and `theme.py` would cover the most testable pure-logic code without requiring GUI or COM mocking.

---

*Concerns audit: 2026-04-24*
