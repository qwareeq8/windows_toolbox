# Technology Stack

**Analysis Date:** 2026-04-24

## Languages

**Primary:**
- Python 3 - Backend logic, window management, COM automation, system tray, application entry point (`main.py`, `workers.py`, `explorer_columns.py`, `bridge.py`, `webview.py`, `settings.py`, `settings_state.py`, `snap_service.py`, `theme.py`, `capture_guard.py`, `startup_shortcut.py`, `app_config.py`)
- JavaScript (JSX) - Frontend UI in React (`frontend/src/main.jsx`, `frontend/src/app.jsx`, `frontend/src/pages.jsx`, `frontend/src/panels.jsx`, `frontend/src/primitives.jsx`, `frontend/src/theme.jsx`, `frontend/src/bridge.js`, `frontend/src/icons.jsx`)

**Secondary:**
- Pascal (Inno Setup) - Installer script (`installer/virelo.iss`)
- PowerShell - Build automation (`scripts/build-installer.ps1`)

## Runtime

**Environment:**
- Python 3 (no `.python-version` file; PySide6>=6.6 implies Python 3.8+)
- Node.js (no `.nvmrc` file; used only for frontend build tooling)
- Target OS: Windows only (enforced at startup via `sys.platform != "win32"` guard in `main.py:17`)
- Requires admin privileges (UAC elevation via `ShellExecuteW` in `main.py:1396-1412`)

**Package Manager:**
- pip (Python) - `requirements.txt`
- npm (JavaScript) - `frontend/package.json`
- Lockfile: `frontend/package-lock.json` present

## Frameworks

**Core:**
- PySide6 >=6.6 - Qt6 bindings for GUI shell, system tray, QWebEngineView (embedded Chromium), QWebChannel (Python<->JS bridge), QSettings (persistent settings), QThread (background workers)
- React 19.1.0 - Frontend UI rendered inside QWebEngineView
- Vite 6.3.4 - Frontend build tool and dev server

**Testing:**
- Not detected - No test framework configured (`tests/` directory absent, no test runner in `package.json` or `requirements.txt`)

**Build/Dev:**
- PyInstaller >=6.0 - Bundles Python app into `Virelo.exe` (spec: `Windows Toolbox.spec`)
- pyinstaller-hooks-contrib >=2024.6 - Additional PyInstaller hooks for dependencies
- @vitejs/plugin-react 4.5.2 - Vite plugin for JSX/React compilation
- Inno Setup 6 - Windows installer builder (`installer/virelo.iss`)

## Key Dependencies

**Critical (Python):**
- `PySide6` >=6.6 - Entire application shell: QMainWindow, QSystemTrayIcon, QWebEngineView, QWebChannel, QThread, QSettings, QTimer, signals/slots. This is the backbone framework.
- `keyboard` >=0.13.5 - Global hotkey capture for snap/restore key detection (press hooks in `main.py:626-627`, key capture sessions in `workers.py:79-141`)
- `pywin32` >=306 - Windows API access: `win32gui`, `win32api`, `win32con`, `win32event`, `win32com.client`, `pythoncom`, `pywintypes`, `winerror`. Used for window manipulation, monitor enumeration, Explorer COM automation, mutex, startup shortcuts.
- `comtypes` >=1.3.0 - Low-level COM interface access for Explorer column auto-sizing via `IColumnManager`, `IServiceProvider`, `IFolderView2` (`explorer_columns.py`)

**Critical (JavaScript):**
- `react` ^19.1.0 - UI component framework
- `react-dom` ^19.1.0 - React DOM renderer

**Infrastructure:**
- `ctypes` (stdlib) - Direct Win32 API calls: `user32.dll`, `kernel32.dll`, `shell32.dll`, `shcore.dll`, `dwmapi.dll` (`main.py:165-166`, `explorer_columns.py`)
- `winreg` (stdlib) - Windows registry reading for system theme detection (`theme.py:26-33`)

## Configuration

**Application Settings:**
- Persisted via `QSettings` (Windows Registry under `HKCU\Software\Yusuf Qwareeq\Virelo\Settings`)
- Settings class: `settings.py` - reads/writes 11 keys from QSettings
- SettingsState class: `settings_state.py` - JSON-friendly wrapper with validation, coercion, and partial updates
- Defaults defined in `app_config.py:8-20` (DEFAULTS dict)
- 11 settings keys: `snap_key`, `restore_key`, `enable_snap`, `snap_presses`, `snap_interval`, `width_pct`, `height_pct`, `ex_auto_size`, `game_mode_enabled`, `run_at_startup`, `theme`

**Environment Variables:**
- `LOCALAPPDATA` - Log file location (`main.py:64`)
- `APPDATA` - Startup shortcut path (`main.py:569`)
- `VIRELO_DEV` - Force dev mode for WebView (connect to Vite dev server) (`webview.py:52`)

**Logging:**
- RotatingFileHandler at `%LOCALAPPDATA%\Virelo\virelo.log` (512KB max, 5 backups) (`main.py:62-94`)
- Crash log at `%LOCALAPPDATA%\Virelo\crash.log` via `faulthandler` (`main.py:116-127`)
- Console handler for development (INFO level) (`main.py:96-108`)
- JS console messages routed to Python logging via `VireloWebPage.javaScriptConsoleMessage` (`webview.py:79-86`)

**Build:**
- `Windows Toolbox.spec` - PyInstaller spec: entry point `main.py`, bundles `icon.ico` and `frontend/dist/`, UPX compression, windowed mode (no console)
- `frontend/vite.config.js` - Vite config: base `./` (relative paths for file:// loading), all assets inlined (100KB limit), output to `frontend/dist/`, dev server on port 5173
- `installer/virelo.iss` - Inno Setup: version 1.4.2, LZMA2 compression, x64 only, installs to `%ProgramFiles%\Virelo`
- `scripts/build-installer.ps1` - Orchestrates PyInstaller build then Inno Setup compilation

## Architecture Pattern

**Hybrid Desktop App:**
- Python backend (PySide6/Qt) hosts an embedded Chromium browser (QWebEngineView)
- React frontend renders all UI inside the WebView
- Communication via QWebChannel: Python `VireloBridge` QObject registered as `"bridge"`, JS accesses via `channel.objects.bridge`
- Dev mode: React served from Vite dev server at `localhost:5173`
- Release mode: React built to `frontend/dist/`, loaded via `file://` URL

**Singleton Instance:**
- Win32 named mutex `Global\Virelo_Mutex` prevents multiple instances (`main.py:1425-1427`)

## Platform Requirements

**Development:**
- Windows 10/11 (x64)
- Python 3.8+ with pip
- Node.js with npm (for frontend builds)
- Admin privileges recommended (app auto-elevates)
- Vite dev server on port 5173 for frontend hot-reload

**Production:**
- Windows 10/11 (x64) - enforced by Inno Setup `ArchitecturesAllowed=x64os`
- System tray support required (`main.py:1433-1437`)
- Admin privileges required (auto-elevation at startup)
- No external runtime needed (PyInstaller bundles everything)

**Build Pipeline:**
1. `cd frontend && npm run build` - Build React frontend to `frontend/dist/`
2. `python -m PyInstaller --clean --noconfirm "Windows Toolbox.spec"` - Bundle into `dist/Virelo/`
3. `ISCC.exe installer/virelo.iss` - Create `installer/dist/VireloSetup.exe`

---

*Stack analysis: 2026-04-24*
