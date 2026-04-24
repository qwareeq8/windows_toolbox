# Coding Conventions

**Analysis Date:** 2026-04-24

## Naming Patterns

**Files (Python):**
- `snake_case.py` for all modules: `main.py`, `bridge.py`, `settings_state.py`, `explorer_columns.py`, `snap_service.py`, `capture_guard.py`, `startup_shortcut.py`, `app_config.py`
- Flat layout -- all Python modules live at project root (no `src/` directory)

**Files (JavaScript/JSX):**
- `lowercase.jsx` or `lowercase.js` for all frontend files: `app.jsx`, `pages.jsx`, `panels.jsx`, `primitives.jsx`, `theme.jsx`, `bridge.js`, `icons.jsx`, `main.jsx`
- All source in `frontend/src/`

**Functions (Python):**
- `snake_case` for all functions and methods: `get_monitor_rect()`, `_is_window_fullscreen()`, `autosize_explorer_columns()`
- Private methods and functions prefixed with `_`: `_apply_side_effects()`, `_get_window_dwm_rect()`, `_canonicalize_path()`
- Module-level helper functions are private by convention: `_safe_int()`, `_safe_bool()` in `settings.py`

**Functions (JavaScript):**
- `camelCase` for all functions: `bridgeToState()`, `stateToBridge()`, `handleSave()`, `handleDiscard()`
- React components use `PascalCase`: `TitleBar`, `NavItem`, `Sidebar`, `Footer`, `MonitorPreview`, `SnapPage`, `CommandPalette`
- Hook-style functions: `useTokens()`, `useTheme()`, `useBridgeSync()`

**Variables (Python):**
- `snake_case` for all variables: `snap_enabled`, `theme_mode`, `press_times`
- Private instance attributes prefixed with `_`: `self._state`, `self._snap`, `self._main_window`
- Module-level constants use `UPPER_SNAKE_CASE`: `APP_NAME`, `FULLSCREEN_TOLERANCE`, `MIN_AUTOSIZE_INTERVAL_PER_TAB_MS`, `TRANSIENT_HRESULT_CODES`
- Win32 API constants use `UPPER_SNAKE_CASE`: `LVM_FIRST`, `WM_KEYDOWN`, `HDM_FIRST`

**Variables (JavaScript):**
- `camelCase` for all variables and state: `snapEnabled`, `pressCount`, `statusMsg`
- React state setter shorthand: `set` (a merged setter function, not individual `setX` per field)
- Constants use `UPPER_SNAKE_CASE`: `MOCK_SETTINGS`, `MOCK_BRIDGE`, `ACCENTS`, `LIGHT`, `DARK`, `DENSITIES`

**Classes (Python):**
- `PascalCase`: `MainWindow`, `ShiftSnapRestore`, `VireloBridge`, `SettingsState`, `CaptureGuard`, `SnapService`
- `PascalCase` for dataclasses: `TabAutosizeState`, `DedupeKey`, `DedupeEntry`, `ShellWindow`, `TabState`
- COM interface classes: `IServiceProvider`, `IColumnManager` (matching COM naming)

**Types (Python):**
- Type annotations from `typing` used on function signatures: `Optional`, `Tuple`, `Dict`, `Deque`, `Union`, `Set`, `Callable`
- Return type annotations on public methods in `bridge.py`: `def get_settings(self) -> str:`
- Dataclass field types annotated: `tab_id: int`, `path: str`, `view_mode: Optional[int]`
- No strict enforcement -- many functions in `main.py` lack annotations

## Code Style

**Formatting (Python):**
- No formatter config file detected (no `pyproject.toml`, `setup.cfg`, `ruff.toml`, `.flake8`)
- De facto style: 4-space indentation, ~88-100 char line length (not strictly enforced)
- Trailing commas used in multi-line structures
- Parenthesized line continuations preferred over backslash

**Formatting (JavaScript):**
- No formatter config file detected (no `.prettierrc`, `.eslintrc`, `biome.json`)
- Single quotes for strings
- 2-space indentation in JSX
- Inline styles used everywhere (no CSS files, no CSS-in-JS library)
- Compact one-liner style for simple components and style objects

**Linting:**
- No linting tools configured for either Python or JavaScript
- One `// eslint-disable-next-line no-undef` comment in `frontend/src/bridge.js` line 51 (for `QWebChannel` global)

## Import Organization

**Python import order:**
1. Standard library: `import ctypes`, `import logging`, `import os`, `import sys`, `import threading`, `import time`
2. Third-party: `import keyboard`, `import win32api`, `import win32con`, `from PySide6 import QtCore, QtGui, QtWidgets`
3. Local modules: `from app_config import ...`, `from bridge import VireloBridge`, `from settings import Settings`
- No blank line enforced between groups (varies by file)
- `from __future__ import annotations` used in `explorer_columns.py` only

**JavaScript import order:**
1. React: `import React from 'react'`
2. Local modules: `import { useTokens, useTheme } from './theme.jsx'`
- Explicit `.jsx` / `.js` extensions on all local imports
- Named imports preferred: `import { ThemeProvider, useTheme, useTokens } from './theme.jsx'`
- Default exports used for main app component only: `export default function VireloApp`

**Path Aliases:**
- None -- all imports use relative paths (Python) or `./` relative paths (JavaScript)

## Error Handling

**Python -- Defensive exception swallowing:**
- The dominant pattern is `try/except Exception: pass` for non-critical operations
- Used extensively for Win32 API calls that may fail transiently:
  ```python
  # Pattern from main.py _get_window_dwm_rect()
  try:
      result = ctypes.windll.dwmapi.DwmGetWindowAttribute(...)
      if result == 0:
          return (rect.left, rect.top, rect.right, rect.bottom)
  except Exception:
      pass
  ```

**Python -- LOG.exception for important failures:**
- Bridge methods log and return JSON error objects:
  ```python
  # Pattern from bridge.py
  except Exception as e:
      LOG.exception("get_settings failed")
      return json.dumps({"error": str(e)})
  ```

**Python -- JSON result objects for bridge communication:**
- Success: `{"ok": True, "applied": {...}}` or `{"ok": True, "message": "..."}`
- Failure: `{"ok": False, "error": "descriptive message"}`
- All `VireloBridge` `@Slot` methods return JSON strings, even on error
- Input validation before processing: `if target not in ("snap", "restore"): return json.dumps({"ok": False, ...})`

**Python -- Safe type coercion with defaults:**
- `_safe_int(val, default)` and `_safe_bool(val, default)` in `settings.py` -- never raise, always return a usable value
- `normalize_theme_mode(value, default)` and `normalize_snap_presses(value)` in `app_config.py` and `theme.py`

**Python -- COM error handling:**
- Transient COM errors identified by HRESULT code set: `TRANSIENT_HRESULT_CODES` in `explorer_columns.py`
- Circuit breaker pattern with exponential backoff for persistent failures
- `(success: bool, method: str, is_transient: bool)` return tuples distinguish recoverable from permanent errors

**JavaScript -- try/catch around JSON.parse:**
- Every bridge callback wraps `JSON.parse` in try/catch:
  ```javascript
  // Pattern from app.jsx
  bridge.get_settings((json) => {
    try {
      const settings = JSON.parse(json);
      setState(bridgeToState(settings));
    } catch (e) {
      console.error('[app] Failed to parse initial settings:', e);
    }
  });
  ```

**JavaScript -- Prefix-tagged console logging:**
- `console.error('[app] ...')`, `console.error('[bridge] ...')` for categorized JS console output

## Logging

**Framework:** Python `logging` module with named logger `"Virelo"`

**Logger initialization pattern:**
```python
# Every module that logs:
LOG = logging.getLogger("Virelo")
```

**Log levels used:**
- `LOG.debug(...)` -- COM enumeration, poll cycles, rate limiting details
- `LOG.info(...)` -- Autosize success, worker start/stop, theme changes
- `LOG.warning(...)` -- COM cache corruption, thread shutdown timeouts, circuit breaker trips
- `LOG.error(...)` -- JS console errors routed to Python, critical failures
- `LOG.exception(...)` -- With full traceback for unexpected exceptions in bridge/snap operations

**Format:**
```
%(asctime)s [%(levelname)s] %(message)s
```

**Destinations:**
- `RotatingFileHandler`: 512KB max, 5 backups, UTF-8, saved to `%LOCALAPPDATA%/Virelo/virelo.log`
- `StreamHandler` (console): INFO and above for development feedback
- Crash diagnostics: `faulthandler` enabled, writes to `crash.log` in the same log directory

**JS console routing:** `VireloWebPage.javaScriptConsoleMessage()` in `webview.py` routes JS `console.log/warn/error` to Python `LOG.debug/warning/error`

## Comments

**When to Comment:**
- Module-level docstrings on all files describing purpose: `"""QWebChannel bridge between React frontend and Python backend..."""`
- Section separators using `# --- heading ---` or `# ------...\n# heading\n# ------...` banners in `main.py`
- Inline comments for Win32 constant definitions: `LVM_FIRST = 0x1000  # ...`
- COM interface explanations and HRESULT codes documented inline in `explorer_columns.py`
- "Why" comments for non-obvious logic: `# Fallback: try relative to script directory`

**Docstrings:**
- Multi-line docstrings use Google-style Args/Returns blocks on complex functions:
  ```python
  def get_monitor_rect(hwnd: int, use_work_area: bool = True) -> Optional[Tuple[int, int, int, int]]:
      """
      Get monitor rectangle for a given window.
      
      Args:
          hwnd: Window handle
          use_work_area: If True, return work area (taskbar-adjusted).
      
      Returns:
          (left, top, right, bottom) tuple or None
      """
  ```
- Single-line docstrings for simple methods: `"""Return all settings as a JSON string."""`
- Class docstrings describe purpose, registration, and public API

**JSDoc:**
- Block comments with `/** ... */` used sparingly, mainly in `bridge.js` and `main.jsx`
- Single-line `//` comments for section headers in JSX files: `// Pages for v2: Window Snap, Explorer, ...`

## Function Design

**Size:**
- Small utility functions preferred: `_safe_int`, `_safe_bool`, `normalize_theme_mode`, `resolve_theme` are 3-10 lines
- Service classes wrap complex logic behind narrow APIs: `SnapService` wraps `ShiftSnapRestore`
- Large methods exist in `main.py` (`MainWindow.__init__` ~90 lines, `_snap` ~90 lines) and `workers.py` (`ExplorerAutosizeEngine.step` ~250 lines)

**Parameters:**
- Keyword arguments with defaults for optional behavior: `use_work_area: bool = True`, `caller_owns_com: bool = False`
- Dependency injection for testability: `read_registry=None` in `get_windows_theme()`, `exists=os.path.exists` in `startup_shortcut.py`
- Callable injection for workers: `autosize_try: Callable[[int, Optional[str]], Tuple[bool, str, bool]]`

**Return Values:**
- `Optional[T]` for functions that may fail: `-> Optional[Tuple[int, int, int, int]]`
- `dict` with `ok`/`error` keys for bridge operations
- `(success, method, is_transient)` tuples for autosize operations
- `bool` for simple checks: `is_admin()`, `_is_window_interactive()`

## Module Design

**Exports (Python):**
- No `__all__` declarations
- Classes and functions at module level -- no nested namespaces
- Related functionality grouped in single files rather than packages

**Exports (JavaScript):**
- Named exports: `export { Toggle, Button, Card, Row, ... }` in `primitives.jsx`
- Named + default mixed: `export default function VireloApp` in `app.jsx` with named exports from other files
- Re-exports from `theme.jsx`: `export { ThemeProvider, useTheme, useTokens, ACCENTS, LIGHT, DARK, DENSITIES, FONT, MONO }`

**Barrel Files:**
- None used -- every import references a specific file directly

## Architecture Patterns

**Bridge pattern (Python-JS communication):**
- `VireloBridge` (`bridge.py`) is a `QObject` with `@Slot` methods and `Signal` attributes
- All data crosses the bridge as JSON strings -- no direct Python object exposure
- `SettingsState` (`settings_state.py`) provides JSON-friendly read/write over `Settings`
- `SnapService` (`snap_service.py`) wraps `ShiftSnapRestore` for narrow API surface

**State management (Frontend):**
- Single `useState` object in `VireloApp` holds all settings state
- Merged setter: `const set = (p) => { setState((s) => ({ ...s, ...p })); setUnsaved(true); }`
- Theme state managed separately via `ThemeProvider` context
- `bridgeToState()` / `stateToBridge()` map between Python snake_case keys and React camelCase keys

**Inline styles (Frontend):**
- All styling is inline via `style={{...}}` objects
- No CSS files, no CSS modules, no styled-components, no Tailwind
- Theme tokens consumed via `useTokens()` hook: `const t = useTokens();`
- Hover state managed per-component via `useState(false)`

**Thread model (Python):**
- Qt `QThread` + `moveToThread` pattern for background workers
- `KeyCaptureWorker` and `ExplorerAutosizeWorker` are `QObject` subclasses with `@Slot` run methods
- Workers communicate back via Qt Signals: `captured`, `cancelled`, `finished`
- `CaptureGuard` uses `threading.Lock` for mutual exclusion of key capture sessions

---

*Convention analysis: 2026-04-24*
