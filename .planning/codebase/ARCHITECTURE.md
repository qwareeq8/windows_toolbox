# Architecture

**Analysis Date:** 2026-04-24

## Pattern Overview

**Overall:** Hybrid Desktop App -- PySide6 (Qt) shell hosting a React SPA via QWebEngineView, communicating through a typed QWebChannel bridge.

**Key Characteristics:**
- Python backend owns all OS-level logic (window management, keyboard hooks, COM automation, system tray)
- React frontend owns all UI rendering (settings pages, theme, command palette)
- A single `VireloBridge` QObject is the sole communication channel between the two layers
- Background work runs on dedicated QThreads with QObject workers (`moveToThread` pattern)
- Settings flow: QSettings (Windows registry) -> `Settings` -> `SettingsState` -> JSON -> Bridge -> React state
- App requires Windows admin elevation; enforced at startup via `ShellExecuteW("runas")`

## Layers

**1. Entry Point / Application Shell (`main.py` lines 1396-1451):**
- Purpose: Bootstrap the QApplication, enforce single-instance mutex, request admin elevation, create `MainWindow`
- Location: `main.py` -- `main()` function
- Contains: DPI awareness setup, Qt attribute configuration, `win32event.CreateMutex` for single-instance guard
- Depends on: PySide6, win32api, ctypes
- Used by: Nothing (top-level entry)

**2. Main Window (`main.py` lines 905-1389):**
- Purpose: Orchestrates all subsystems -- tray icon, bridge wiring, theme sync, background threads, frameless window resize
- Location: `main.py` -- `MainWindow(QtWidgets.QMainWindow)`
- Contains: System tray menu, keyboard shortcuts (Ctrl+T, Ctrl+Enter, F1), theme timer, key capture thread management, explorer autosize thread management
- Depends on: `Settings`, `SettingsState`, `SnapService`, `VireloBridge`, `VireloWebView`, `ShiftSnapRestore`, `CaptureGuard`, workers
- Used by: `main()` entry point

**3. Bridge Layer:**
- Purpose: Narrow, validated JSON-based communication between Python backend and React frontend
- Location: `bridge.py` -- `VireloBridge(QObject)`
- Contains: Qt Slots (callable from JS) and Signals (push to JS) for settings CRUD, snap actions, key capture, theme control, startup toggle
- Depends on: `SettingsState`, `SnapService`, `MainWindow` (set post-construction)
- Used by: React frontend via `channel.objects.bridge`

**4. Settings Layer:**
- Purpose: Persistent settings via QSettings (Windows registry) with JSON serialization for the bridge
- Location: `settings.py` -- `Settings`, `settings_state.py` -- `SettingsState`, `app_config.py` -- defaults/constants
- Contains: 11 setting keys (snap_key, restore_key, enable_snap, snap_presses, snap_interval, width_pct, height_pct, ex_auto_size, game_mode_enabled, run_at_startup, theme)
- `Settings` reads/writes QSettings directly; `SettingsState` wraps it with type coercion, validation, and JSON output
- Depends on: PySide6.QtCore.QSettings, `app_config.DEFAULTS`
- Used by: `MainWindow`, `VireloBridge`, `ShiftSnapRestore`

**5. Snap/Restore Engine (`main.py` lines 609-895):**
- Purpose: Detect keyboard multi-press events and snap/restore the foreground window
- Location: `main.py` -- `ShiftSnapRestore(QtCore.QObject)`
- Contains: Keyboard hook management (`keyboard` library), press timing with deque, window position save/restore, fullscreen/game detection via Win32 API, DWM extended frame bounds
- Depends on: `keyboard`, `win32gui`, `win32con`, `ctypes`, `Settings`
- Used by: `MainWindow` (owned), `SnapService` (wrapped)

**6. Explorer Column Autosize:**
- Purpose: Automatically resize File Explorer Detail view columns when navigating to a new folder
- Location: `explorer_columns.py` -- COM IColumnManager interface, `workers.py` -- `ExplorerAutosizeEngine` (tab-aware state machine), `ExplorerAutosizeWorker` (QThread worker)
- Contains: COM interface definitions (IServiceProvider, IColumnManager), Windows 11 tab tracking, debounce/settle/rate-limit/circuit-breaker logic, deduplication with TTL
- Depends on: `comtypes`, `pythoncom`, `win32com.client`, `Shell.Application` COM object
- Used by: `MainWindow` (starts/stops worker thread)

**7. Background Workers (`workers.py`):**
- Purpose: Long-running tasks that must not block the Qt event loop
- Location: `workers.py`
- Contains:
  - `KeyCaptureWorker` (lines 638-663): Captures a single keypress for rebinding snap/restore keys. Uses `KeyCaptureSession` internally.
  - `ExplorerAutosizeWorker` (lines 666-978): Polls Explorer windows via COM and delegates to `ExplorerAutosizeEngine` for tab-aware column auto-sizing.
  - `KeyCaptureSession` (lines 78-141): Pure-Python key capture with timeout, cancel, and `keyboard.hook`.
  - `ExplorerAutosizeEngine` (lines 174-584): Stateful engine with per-tab tracking, deduplication, circuit breakers, and exponential backoff.
- Depends on: `keyboard`, `pythoncom`, `win32com.client`, PySide6.QtCore
- Used by: `MainWindow` (via `QThread.moveToThread` pattern)

**8. Frontend React App:**
- Purpose: Render the entire settings UI, handle user interactions, relay changes to Python via bridge
- Location: `frontend/src/`
- Contains: App shell (`app.jsx`), 5 page components (`pages.jsx`), command palette (`panels.jsx`), design system primitives (`primitives.jsx`), theme system (`theme.jsx`), SVG icons (`icons.jsx`), bridge client (`bridge.js`)
- Depends on: React 19, Vite, QWebChannel
- Used by: `VireloWebView` (loads as webpage)

**9. WebView Host (`webview.py`):**
- Purpose: Create and configure the QWebEngineView that hosts the React frontend
- Location: `webview.py` -- `VireloWebView(QWebEngineView)`, `VireloWebPage(QWebEnginePage)`
- Contains: Dev mode detection (Vite dev server on localhost:5173 vs built files), QWebChannel setup, JS console routing to Python logging
- Depends on: PySide6.QtWebEngineWidgets, `VireloBridge`
- Used by: `MainWindow`

**10. Support Modules:**
- `theme.py`: Pure functions for theme resolution (system/dark/light), Windows registry read for system theme
- `capture_guard.py`: Thread-safe boolean guard preventing concurrent key captures
- `snap_service.py`: Thin facade over `ShiftSnapRestore` for the bridge layer
- `startup_shortcut.py`: Logic for creating/selecting the correct Python executable for startup shortcuts

## Data Flow

**Settings Read (App Startup):**

1. `Settings.__init__()` reads all 11 keys from `QSettings` (Windows registry) with type coercion and defaults from `app_config.DEFAULTS`
2. `MainWindow.__init__()` creates `SettingsState(settings)` wrapper
3. `VireloBridge.__init__()` receives `SettingsState` reference
4. React `VireloApp` calls `bridge.get_settings()` on mount
5. Bridge Slot `get_settings()` calls `SettingsState.get_json()` -> returns JSON string
6. React parses JSON, maps Python keys to React state keys via `bridgeToState()`

**Settings Write (User Saves):**

1. User modifies controls in React pages -> `app.set({key: value})` updates local React state, marks `unsaved=true`
2. User clicks "Save changes" -> `bridge.save_settings(stateToBridge(state))` sends JSON to Python
3. Bridge Slot `save_settings()` calls `SettingsState.apply_partial(data)` which validates, coerces types, writes to `Settings`, calls `Settings.save()` (QSettings)
4. On success, bridge emits `settings_changed` Signal with full updated JSON
5. Bridge calls `_apply_side_effects()` which updates `MainWindow` state: snap enabled, explorer thread, key bindings, theme

**Snap Trigger (Keyboard):**

1. `keyboard.on_press_key()` callback fires `ShiftSnapRestore._on_press()`
2. Press recorded in timing deque; if count reaches threshold within interval, `triggered` Signal emits
3. Signal connected to `ShiftSnapRestore.perform(restore=bool)` Slot
4. `perform()` gets foreground window via `USER32.GetForegroundWindow()`, checks fullscreen/game mode, calculates target position from settings percentages and monitor work area, calls `USER32.MoveWindow()`

**Explorer Column Autosize:**

1. `ExplorerAutosizeWorker.run()` initializes COM (STA), caches `Shell.Application`
2. Main loop: `iter_tabs()` enumerates Explorer windows via COM `Shell.Application.Windows()`, extracts (hwnd, tab_id, path, view_mode) per tab
3. `ExplorerAutosizeEngine.step()` processes each tab: debounce -> settle -> rate-limit check -> deduplication check -> window interactivity check -> attempt autosize
4. Autosize calls `explorer_columns.autosize_explorer_columns()` which obtains `IColumnManager` via COM `IServiceProvider/SID_SFolderView` and sets `CM_WIDTH_AUTOSIZE` on all visible columns
5. On failure: exponential backoff with circuit breaker (5 failures -> 5s cooldown)

**Theme Sync:**

1. `MainWindow._apply_theme_mode(mode)` stores mode, starts/stops 2-second polling timer
2. If mode is "system": `_sync_system_theme()` reads Windows registry `AppsUseLightTheme`, calls `resolve_theme()`, emits `bridge.theme_applied` Signal
3. React `AppWithBridge` receives Signal, updates `tweaks.theme` in `ThemeProvider`
4. All components re-render with new theme tokens from `useTokens()`

**State Management:**
- Python side: `Settings` object holds canonical state, persisted to Windows registry via `QSettings`. `SettingsState` is the read/write facade.
- React side: `useState` in `VireloApp` holds working copy. Bridge Signals push external changes. `bridgeToState()` / `stateToBridge()` handle key mapping between Python snake_case and React camelCase.

## Key Abstractions

**VireloBridge (`bridge.py`):**
- Purpose: Single point of contact between Python and JavaScript
- Pattern: Qt Slots for JS->Python calls (request/response via callbacks), Qt Signals for Python->JS pushes
- All data exchanged as JSON strings; no Python objects exposed
- Signals: `settings_changed(str)`, `theme_applied(str)`, `snap_status(str, int)`, `capture_status(str)`

**ExplorerAutosizeEngine (`workers.py` lines 174-584):**
- Purpose: Stateful tab-aware autosize state machine
- Pattern: Step-based engine called in a polling loop. Per-tab state includes navigation tracking, debounce timers, retry scheduling, circuit breakers, deduplication cache with TTL.
- Key data structure: `TabAutosizeState` dataclass per (hwnd, path) tuple

**ShiftSnapRestore (`main.py` lines 609-895):**
- Purpose: Multi-press keyboard trigger with snap/restore actions
- Pattern: Keyboard hook -> timing deque -> Qt Signal -> Win32 window manipulation
- Stores original window positions in `_orig_sizes` dict for restore capability

**Theme System (`frontend/src/theme.jsx` + `theme.py`):**
- Purpose: Consistent dark/light/system theming across the entire UI
- Pattern: React Context (`ThemeProvider`) exposes `useTokens()` hook returning a flat token object (colors, spacing, typography). Python side reads Windows registry for system theme, pushes changes via bridge Signal.
- Tweakable knobs: theme (dark/light), accent color (slate/teal/blue/rust/purple), density (compact/cozy/comfortable), border radius, sidebar mode

## Entry Points

**Application Entry (`main.py` line 1450):**
- Location: `main.py` -- `if __name__ == "__main__": main()`
- Triggers: User runs `python main.py` or launches `Virelo.exe` (PyInstaller frozen)
- Responsibilities: Admin elevation, single-instance check, QApplication creation, `MainWindow` construction, event loop

**Frontend Entry (`frontend/src/main.jsx`):**
- Location: `frontend/src/main.jsx`
- Triggers: Loaded by QWebEngineView (from Vite dev server or `frontend/dist/index.html`)
- Responsibilities: Initialize bridge connection, resolve initial theme, render `Root` -> `AppWithBridge` -> `VireloApp`

**Bridge Entry (`frontend/src/bridge.js`):**
- Location: `frontend/src/bridge.js` -- `getBridge()`
- Triggers: Called by `main.jsx` on mount
- Responsibilities: Connect to QWebChannel, return `channel.objects.bridge` or mock bridge for dev mode

## Error Handling

**Strategy:** Defensive try/except at every boundary; log exceptions, return error JSON to frontend, never crash the main loop.

**Patterns:**
- Bridge Slots wrap all logic in try/except, return `{"ok": false, "error": "..."}` JSON on failure, log via `LOG.exception()`
- Background workers catch exceptions per-iteration, log, and continue with a fallback sleep interval
- COM errors in Explorer autosize classified as transient vs non-transient; transient errors trigger retry with backoff, non-transient errors stop retries
- `faulthandler` enabled at startup writing to `crash.log` for segfault diagnostics
- Win32 API calls wrapped individually with exception handling; failures return None/False rather than propagating

## Cross-Cutting Concerns

**Logging:**
- Single rotating file logger named "Virelo" (`RotatingFileHandler`, 512KB, 5 backups)
- Log location: `%LOCALAPPDATA%/Virelo/virelo.log`
- Crash log: `%LOCALAPPDATA%/Virelo/crash.log` (faulthandler)
- Console handler at INFO level for development
- JS `console.*` calls routed to Python logger via `VireloWebPage.javaScriptConsoleMessage()`

**Validation:**
- `SettingsState.apply_partial()` validates types and ranges for all 11 setting keys before writing
- Bridge Slots validate input parameters (e.g., `target in ("snap", "restore")`, `mode in ("system", "dark", "light")`)
- `_safe_int()` and `_safe_bool()` in `settings.py` guard against corrupt QSettings values

**Authentication/Authorization:**
- App requires Windows admin privileges; `main()` re-launches with `runas` if not admin
- No user authentication -- single-user desktop utility

**Threading Model:**
- Main thread: Qt event loop, UI, bridge Slots/Signals
- Key capture thread: `QThread` with `KeyCaptureWorker` (short-lived, one per capture)
- Explorer autosize thread: `QThread` with `ExplorerAutosizeWorker` (long-lived, COM STA apartment, message pumping)
- `CaptureGuard`: Thread-safe mutex preventing concurrent key captures
- `ShiftSnapRestore`: Keyboard hooks fire on arbitrary threads; `triggered` Signal marshals to main thread

---

*Architecture analysis: 2026-04-24*
