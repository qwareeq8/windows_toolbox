# External Integrations

**Analysis Date:** 2026-04-24

## APIs & External Services

**None.** Virelo is a fully offline desktop application with no network calls, no cloud APIs, and no telemetry. All functionality operates locally through Windows OS APIs.

## Windows OS API Integrations

Virelo deeply integrates with multiple Windows subsystems. These are the application's core external dependencies.

**Win32 Window Management:**
- Purpose: Snap/restore windows to configured size and position
- APIs used: `user32.MoveWindow`, `user32.GetWindowRect`, `user32.GetForegroundWindow`, `win32gui.GetWindowPlacement`, `win32gui.ShowWindow`, `win32gui.EnumWindows`, `win32gui.IsWindowVisible`, `win32gui.GetWindowLong`, `win32gui.GetClassName`
- Files: `main.py:165-168` (constants), `main.py:756-895` (ShiftSnapRestore._snap/_restore)
- Package: `ctypes.windll.user32`, `pywin32` (`win32gui`, `win32api`, `win32con`)

**DWM (Desktop Window Manager):**
- Purpose: Accurate fullscreen detection using extended frame bounds (excludes invisible window borders)
- API: `dwmapi.DwmGetWindowAttribute` with `DWMWA_EXTENDED_FRAME_BOUNDS`
- Files: `main.py:246-272` (_get_window_dwm_rect)
- Package: `ctypes.windll.dwmapi`

**Monitor Enumeration:**
- Purpose: Multi-monitor awareness for window snapping (work area vs full bounds)
- APIs: `win32api.MonitorFromWindow`, `win32api.GetMonitorInfo`
- Files: `main.py:221-243` (get_monitor_rect)
- Package: `pywin32` (`win32api`, `win32con`)

**DPI Awareness:**
- Purpose: High-DPI display support
- APIs: `shcore.SetProcessDpiAwareness(2)`, fallback to `user32.SetProcessDPIAware`
- Files: `main.py:211-218` (_enable_dpi_awareness)
- Package: `ctypes.windll.shcore`

**Global Keyboard Hooks:**
- Purpose: Detect snap key multi-press sequences and restore key modifier globally (even when app is not focused)
- APIs: `keyboard.on_press_key`, `keyboard.on_release_key`, `keyboard.is_pressed`, `keyboard.hook`
- Files: `main.py:626-627` (ShiftSnapRestore.__init__), `workers.py:79-141` (KeyCaptureSession)
- Package: `keyboard`

**Windows Shell COM (Shell.Application):**
- Purpose: Enumerate open Explorer windows and their tabs, resolve folder paths, detect navigation
- APIs: `win32com.client.Dispatch("Shell.Application")`, `shell.Windows()`, `window.LocationURL`, `window.Document.Folder.Self.Path`, `window.Document.CurrentViewMode`
- Files: `workers.py:740-800` (ExplorerAutosizeWorker.run, iter_tabs closure)
- Package: `pywin32` (`win32com.client`, `pythoncom`)
- COM apartment: STA (single-threaded apartment), initialized once per worker thread via `pythoncom.CoInitializeEx(COINIT_APARTMENTTHREADED)` (`workers.py:731`)

**Explorer Column Manager COM (IColumnManager):**
- Purpose: Auto-size Explorer "Details view" columns when navigating to new folders
- APIs: `IServiceProvider.QueryService`, `IColumnManager.SetColumns` with `CM_WIDTH_AUTOSIZE`, `IFolderView2` (view mode detection)
- Files: `explorer_columns.py` (entire file, 861 lines)
- Package: `comtypes`, `comtypes.client`
- Error handling: Transient COM error detection for 9 HRESULT codes (`explorer_columns.py:31-41`), circuit breaker pattern with exponential backoff (`workers.py:14-18`)

**WScript.Shell COM:**
- Purpose: Create/remove Windows startup shortcuts (.lnk files) in the Startup folder
- APIs: `Dispatch("WScript.Shell")`, `wsh.CreateShortcut`
- Files: `main.py:578-601` (create_startup_shortcut, remove_startup_shortcut)
- Package: `pywin32` (`win32com.client.Dispatch`)

**Windows Registry:**
- Purpose: Detect system light/dark theme preference
- API: `winreg.OpenKey(HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")`, read `AppsUseLightTheme`
- Files: `theme.py:23-37` (get_windows_theme)
- Package: `winreg` (stdlib)

**Windows Mutex:**
- Purpose: Single-instance enforcement (prevent multiple Virelo processes)
- API: `win32event.CreateMutex`, `win32api.GetLastError` checking for `ERROR_ALREADY_EXISTS`
- Files: `main.py:1425-1427`
- Package: `pywin32` (`win32event`, `win32api`, `winerror`)

**UAC Elevation:**
- Purpose: Auto-request admin privileges at startup (required for global keyboard hooks and some window operations)
- API: `ctypes.windll.shell32.ShellExecuteW` with `"runas"` verb
- Files: `main.py:1396-1412` (main function)
- Package: `ctypes`

**App User Model ID:**
- Purpose: Proper taskbar grouping and identity for the application
- API: `shell32.SetCurrentProcessExplicitAppUserModelID("com.yusufqwareeq.virelo")`
- Files: `main.py:1416-1418`
- Package: `ctypes`

## Data Storage

**Databases:**
- None. No database is used.

**Application Settings:**
- Storage: Windows Registry via `QSettings` (PySide6)
- Location: `HKCU\Software\Yusuf Qwareeq\Virelo\Settings`
- Client: `PySide6.QtCore.QSettings` wrapped by `settings.py` (Settings class)
- State layer: `settings_state.py` (SettingsState class with JSON serialization and validation)

**Log Files:**
- Location: `%LOCALAPPDATA%\Virelo\virelo.log` (rotating, 512KB x 5 backups)
- Crash log: `%LOCALAPPDATA%\Virelo\crash.log`
- Files: `main.py:62-127`

**File Storage:**
- Local filesystem only (no cloud storage)
- Startup shortcut: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Virelo.lnk`

**Caching:**
- In-memory only:
  - Explorer autosize deduplication cache with 5-minute TTL (`workers.py:20-46`)
  - LRU bounded path cache (500 entries) (`workers.py:200-201`)
  - Per-tab autosize state with circuit breaker (`workers.py:49-75`)

## Authentication & Identity

**Auth Provider:**
- Not applicable. Virelo is an offline desktop utility. No user accounts, no authentication, no network identity.

## Monitoring & Observability

**Error Tracking:**
- Local rotating log files only (`%LOCALAPPDATA%\Virelo\virelo.log`)
- Python `faulthandler` for crash dumps to `crash.log`
- No remote error tracking service

**Logs:**
- Python `logging` module with `RotatingFileHandler`
- Format: `%(asctime)s [%(levelname)s] %(message)s`
- JS console messages routed to Python log via `VireloWebPage` (`webview.py:79-86`)
- Log levels: DEBUG to file, INFO to console (development)

## CI/CD & Deployment

**Hosting:**
- Not applicable (desktop application distributed as installer)

**CI Pipeline:**
- Not detected (no `.github/workflows/`, no CI config files found)

**Build Pipeline (manual):**
- `scripts/build-installer.ps1` orchestrates: PyInstaller build -> Inno Setup installer
- Output: `installer/dist/VireloSetup.exe`

**Distribution:**
- Windows installer via Inno Setup (`installer/virelo.iss`)
- Version: 1.4.2 (defined in `installer/virelo.iss:5` and `frontend/package.json:4`)
- Architecture: x64 only

## Communication Bridge (Python <-> JavaScript)

**QWebChannel Bridge:**
- Protocol: Qt QWebChannel over WebSocket-like transport (`qt.webChannelTransport`)
- Python side: `VireloBridge` QObject with `@Slot` methods and `Signal` emitters (`bridge.py`)
- JS side: `QWebChannel` client connecting to `channel.objects.bridge` (`frontend/src/bridge.js`)
- Registration: `QWebChannel.registerObject("bridge", bridge_instance)` (`webview.py:115`)

**Signals (Python -> JS):**
- `settings_changed(str)` - Full settings JSON pushed after save
- `theme_applied(str)` - "dark" or "light" effective theme
- `snap_status(str, int)` - Status message with timeout
- `capture_status(str)` - Key capture state: "capturing", "done", "cancelled", "timeout"

**Slots (JS -> Python):**
- `get_settings() -> str` - Read all settings as JSON
- `save_settings(str) -> str` - Partial update with validation
- `reset_defaults() -> str` - Reset to defaults
- `test_snap() -> str` - Trigger test snap
- `capture_key(str) -> str` - Start key capture for "snap" or "restore"
- `apply_theme(str) -> str` - Set theme mode
- `toggle_run_at_startup(bool) -> str` - Toggle startup shortcut
- `get_theme_mode() -> str` - Read current theme
- `get_launch_at_login() -> bool` - Read startup state
- `get_snap_enabled() -> bool` - Read snap enabled state

**Dev Mode Mock:**
- When QWebChannel is unavailable (Vite dev server without Qt host), `frontend/src/bridge.js` provides `MOCK_BRIDGE` with no-op implementations for all slots/signals

## Environment Configuration

**Required env vars:**
- None strictly required (all have fallbacks)

**Optional env vars:**
- `VIRELO_DEV` - Set to "1"/"true"/"yes" to force Vite dev server mode (`webview.py:52`)
- `LOCALAPPDATA` - Log directory (falls back to `~`) (`main.py:64`)
- `APPDATA` - Startup shortcut location (`main.py:569`)
- `ISCC_PATH` - Inno Setup compiler path for build script (`scripts/build-installer.ps1:13`)

**Secrets:**
- None. No API keys, tokens, or credentials. Fully offline application.

## Webhooks & Callbacks

**Incoming:**
- None

**Outgoing:**
- None

---

*Integration audit: 2026-04-24*
