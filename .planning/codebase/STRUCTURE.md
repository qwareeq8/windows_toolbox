# Codebase Structure

**Analysis Date:** 2026-04-24

## Directory Layout

```
Virelo/
├── main.py                  # Application entry point, MainWindow, ShiftSnapRestore engine (~52K)
├── bridge.py                # QWebChannel bridge: VireloBridge QObject (~8.7K)
├── webview.py               # QWebEngineView host for React frontend (~4.7K)
├── workers.py               # Background QThread workers + ExplorerAutosizeEngine (~41K)
├── explorer_columns.py      # COM IColumnManager for Explorer column resizing (~29K)
├── settings.py              # QSettings persistence layer (~3.5K)
├── settings_state.py        # JSON-friendly settings wrapper with validation (~3.4K)
├── app_config.py            # Constants, defaults, app identity (~0.6K)
├── theme.py                 # Theme resolution (system/dark/light) + Windows registry (~1.3K)
├── capture_guard.py         # Thread-safe key capture mutex (~0.4K)
├─��� snap_service.py          # Thin facade over ShiftSnapRestore for bridge (~1.5K)
├── startup_shortcut.py      # Startup shortcut path selection (~0.5K)
├── icon.ico                 # Application icon (Windows ICO format)
├── requirements.txt         # Python dependencies
├── Windows Toolbox.spec     # PyInstaller build spec
│
├── frontend/                # React SPA (Vite + React 19)
│   ├── index.html           # HTML shell with QWebChannel script tag
│   ├── package.json         # npm manifest (virelo-frontend v1.4.2)
���   ├── package-lock.json    # npm lockfile
│   ├── vite.config.js       # Vite config (base: './', single-chunk output)
│   └── src/
│       ├── main.jsx         # React entry: bridge init, Root component, ThemeProvider
│       ├── app.jsx          # App shell: TitleBar, Sidebar, Footer, page routing, state
│       ├── pages.jsx        # 5 page components: Snap, Explorer, Shortcuts, General, About
│       ��── panels.jsx       # Command palette (Ctrl+K)
│       ��── primitives.jsx   # Shared UI primitives: Toggle, Button, Card, Row, Segmented, Stepper, Slider, Kbd, Badge
│       ├── theme.jsx        # ThemeProvider, useTheme, useTokens, color palettes, density scales
│       ├── icons.jsx        # 16 inline SVG icons via Icon component
│       └── bridge.js        # QWebChannel client: getBridge(), mock bridge for dev mode
│
├── branding/                # Installer artwork
│   ├── virelo-icon.svg      # SVG source for app icon
│   ├── installer-header.bmp # Inno Setup header image
│   ├── installer-header_2x.bmp
│   ├── installer-wizard.bmp # Inno Setup wizard image
│   └── installer-wizard_2x.bmp
│
├── installer/               # Inno Setup installer script
│   └── virelo.iss           # Installer definition (AppId, files, shortcuts)
│
├── scripts/                 # Build/utility scripts
│   ├── build-icon.py        # Generate icon.ico from SVG
│   ├── build-installer.ps1  # PowerShell: run Inno Setup compiler
│   └─�� build-installer-bmps.py  # Generate installer BMP artwork
│
└── .planning/               # GSD planning documents
    └���─ codebase/            # Codebase analysis (this directory)
```

## Directory Purposes

**Root (`Virelo/`):**
- Purpose: All Python backend source files live at the root level (flat module layout)
- Contains: 12 `.py` files, 1 `.spec` file, 1 `.ico`, 1 `requirements.txt`
- Key files: `main.py` (entry point + MainWindow + ShiftSnapRestore), `bridge.py` (Python-JS bridge), `workers.py` (background threads + autosize engine)

**`frontend/`:**
- Purpose: Self-contained React SPA built with Vite
- Contains: 7 JSX/JS source files, HTML entry, Vite config, npm manifests
- Key files: `src/app.jsx` (app shell + state management), `src/bridge.js` (QWebChannel client)
- Build output: `frontend/dist/` (generated, not committed -- loaded by QWebEngineView in release mode)

**`branding/`:**
- Purpose: Visual assets for the installer
- Contains: SVG icon source, BMP images for Inno Setup wizard pages (1x and 2x)
- Generated: BMPs created by `scripts/build-installer-bmps.py`

**`installer/`:**
- Purpose: Windows installer definition
- Contains: Single Inno Setup script (`virelo.iss`)
- Depends on: Built PyInstaller output in `dist/Virelo/`

**`scripts/`:**
- Purpose: Build automation utilities
- Contains: Icon generation, installer BMP generation, installer compilation
- Run manually during release process

## Key File Locations

**Entry Points:**
- `main.py`: Python entry -- `main()` at line 1396, `if __name__ == "__main__"` at line 1450
- `frontend/src/main.jsx`: React entry -- creates root, renders `Root` component
- `frontend/index.html`: HTML shell, loads `qrc:///qtwebchannel/qwebchannel.js` and `/src/main.jsx`

**Configuration:**
- `app_config.py`: App identity (APP_NAME, APP_ID, ORGANIZATION), default settings values (`DEFAULTS` dict), log file paths
- `frontend/vite.config.js`: Vite build config (relative base path, single-chunk output, dev server port 5173)
- `frontend/package.json`: npm dependencies (React 19, Vite 6)
- `requirements.txt`: Python dependencies (PySide6, keyboard, pyinstaller, pywin32, comtypes)
- `Windows Toolbox.spec`: PyInstaller spec (bundles `icon.ico` and `frontend/dist/`)

**Core Logic (Python):**
- `main.py` lines 609-895: `ShiftSnapRestore` -- keyboard hook, press timing, snap/restore via Win32
- `main.py` lines 905-1389: `MainWindow` -- tray, theme sync, thread management, frameless resize
- `bridge.py`: `VireloBridge` -- all 12 Slots (get_settings, save_settings, reset_defaults, test_snap, capture_key, apply_theme, get_theme_mode, toggle_run_at_startup, get_launch_at_login, get_snap_enabled) + 4 Signals
- `workers.py` lines 174-584: `ExplorerAutosizeEngine` -- tab-aware state machine
- `workers.py` lines 666-978: `ExplorerAutosizeWorker` -- COM lifecycle, Shell.Application caching, polling loop
- `explorer_columns.py`: COM IColumnManager interface for Explorer column sizing

**Core Logic (Frontend):**
- `frontend/src/app.jsx`: `VireloApp` -- central state, bridge wiring, save/discard/reset handlers, `bridgeToState()`/`stateToBridge()` key mappers
- `frontend/src/pages.jsx`: `SnapPage`, `ExplorerPage`, `ShortcutsPage`, `GeneralPage`, `AboutPage`, `Pg` layout wrapper
- `frontend/src/panels.jsx`: `CommandPalette` -- search/filter/execute commands
- `frontend/src/theme.jsx`: `ThemeProvider`, token definitions (LIGHT/DARK palettes, ACCENTS, DENSITIES)
- `frontend/src/primitives.jsx`: `Toggle`, `Button`, `Card`, `Row`, `Segmented`, `Stepper`, `Slider`, `Kbd`, `Badge`

**Settings Pipeline:**
- `app_config.py`: `DEFAULTS` dict with all 11 setting keys and default values
- `settings.py`: `Settings` class -- reads/writes QSettings (Windows registry), `_safe_int()`, `_safe_bool()` helpers
- `settings_state.py`: `SettingsState` class -- `KEYS` dict with type coercers and validation ranges, `get_json()`, `apply_partial()`, `reset_to_defaults()`

**Support:**
- `theme.py`: `normalize_theme_mode()`, `resolve_theme()`, `toggle_theme_mode()`, `get_windows_theme()` (reads registry `AppsUseLightTheme`)
- `capture_guard.py`: `CaptureGuard` -- thread-safe `try_start()`/`finish()` for exclusive key capture
- `snap_service.py`: `SnapService` -- facade wrapping `ShiftSnapRestore` for bridge consumption
- `startup_shortcut.py`: `startup_shortcut_spec()` -- selects `pythonw.exe` over `python.exe` for startup shortcuts

**Testing:**
- No test files detected in the repository

## Naming Conventions

**Python Files:**
- `snake_case.py` for all modules: `settings_state.py`, `capture_guard.py`, `snap_service.py`, `startup_shortcut.py`, `explorer_columns.py`, `app_config.py`
- Exception: `main.py`, `bridge.py`, `workers.py`, `webview.py`, `theme.py` (single-word names)

**Python Classes:**
- `PascalCase`: `MainWindow`, `ShiftSnapRestore`, `VireloBridge`, `VireloWebView`, `SettingsState`, `CaptureGuard`, `SnapService`, `ExplorerAutosizeEngine`, `ExplorerAutosizeWorker`, `KeyCaptureWorker`, `KeyCaptureSession`, `TabAutosizeState`

**Python Functions:**
- `snake_case` with leading underscore for private: `_init_logger()`, `_enable_dpi_awareness()`, `_is_window_fullscreen()`, `_apply_side_effects()`
- Public functions without underscore: `resource_path()`, `is_admin()`, `get_monitor_rect()`

**Python Constants:**
- `UPPER_SNAKE_CASE`: `APP_NAME`, `DEFAULTS`, `LOG_DIR`, `FULLSCREEN_TOLERANCE`, `MIN_AUTOSIZE_INTERVAL_PER_TAB_MS`

**Frontend Files:**
- `lowercase.jsx` / `lowercase.js`: `app.jsx`, `pages.jsx`, `panels.jsx`, `primitives.jsx`, `theme.jsx`, `icons.jsx`, `bridge.js`, `main.jsx`

**Frontend Components:**
- `PascalCase` function components: `VireloApp`, `TitleBar`, `Sidebar`, `NavItem`, `Footer`, `SnapPage`, `ExplorerPage`, `GeneralPage`, `AboutPage`, `ShortcutsPage`, `CommandPalette`, `MonitorPreview`
- `PascalCase` primitives: `Toggle`, `Button`, `Card`, `Row`, `Segmented`, `Stepper`, `Slider`, `Kbd`, `Badge`

**Frontend Hooks:**
- `useCamelCase`: `useTheme`, `useTokens`, `useBridgeSync`

**Settings Keys:**
- Python side: `snake_case` -- `snap_key`, `restore_key`, `enable_snap`, `snap_presses`, `snap_interval`, `width_pct`, `height_pct`, `ex_auto_size`, `game_mode_enabled`, `run_at_startup`, `theme`
- React side: `camelCase` -- `snapKey`, `restoreKey`, `snapEnabled`, `pressCount`, `interval`, `width`, `height`, `autoSize`, `gameMode`, `launchLogin`, `theme`
- Mapping functions: `bridgeToState()` and `stateToBridge()` in `frontend/src/app.jsx`

## Where to Add New Code

**New Python Feature (e.g., new snap behavior):**
- If it is a new core engine: Add a new `.py` file at root level (e.g., `new_feature.py`)
- If it extends snap/restore: Add methods to `ShiftSnapRestore` class in `main.py` (lines 609-895)
- If it needs bridge access: Add a new `@Slot` method to `VireloBridge` in `bridge.py`, add corresponding Signal if Python needs to push updates
- If it needs background processing: Create a new `QObject` worker class in `workers.py`, manage its `QThread` lifecycle in `MainWindow`
- Wire it in `MainWindow.__init__()` in `main.py`

**New Settings Key:**
1. Add default value to `DEFAULTS` in `app_config.py`
2. Add attribute read in `Settings.__init__()` and write in `Settings.save()` in `settings.py`
3. Add key entry to `SettingsState.KEYS` dict in `settings_state.py` with `(type_coercer, range_or_None)`
4. Add mapping in `bridgeToState()` and `stateToBridge()` in `frontend/src/app.jsx`
5. Add UI control in the appropriate page component in `frontend/src/pages.jsx`
6. If side effects needed: handle in `VireloBridge._apply_side_effects()` in `bridge.py`

**New React Page:**
1. Create the page component in `frontend/src/pages.jsx` following the `Pg` wrapper pattern
2. Add a `NavItem` entry in `Sidebar` in `frontend/src/app.jsx`
3. Add the route key to the `Page` lookup object in `VireloApp` in `frontend/src/app.jsx` (line 291)
4. Add navigation command to `CommandPalette` commands array in `frontend/src/panels.jsx`

**New UI Primitive:**
- Add to `frontend/src/primitives.jsx`, consume `useTokens()` for theming
- Export from the file; import where needed in pages

**New SVG Icon:**
- Add path data to the `paths` object in `frontend/src/icons.jsx`
- Use via `<Icon name="newName" />` -- all icons are 16x16 viewBox, stroke-based

**New Build/Script:**
- Add to `scripts/` directory

## Special Directories

**`frontend/dist/` (not present in repo):**
- Purpose: Vite build output (static HTML/JS/CSS)
- Generated: Yes, by `npm run build` in `frontend/`
- Committed: No (not in repo)
- Used at runtime: Yes -- QWebEngineView loads `frontend/dist/index.html` in release mode (frozen PyInstaller builds)
- Bundled: Yes -- included in PyInstaller via `Windows Toolbox.spec` datas

**`dist/` (not present in repo):**
- Purpose: PyInstaller output directory
- Generated: Yes, by running PyInstaller with `Windows Toolbox.spec`
- Committed: No
- Contains: `dist/Virelo/` folder with frozen executable and all dependencies

**`installer/dist/` (not present in repo):**
- Purpose: Inno Setup output
- Generated: Yes, by `scripts/build-installer.ps1`
- Committed: No
- Contains: `VireloSetup.exe` Windows installer

**`.planning/codebase/`:**
- Purpose: GSD codebase analysis documents
- Generated: Yes, by codebase mapping
- Committed: Yes

## Build Pipeline

**Development:**
1. Run `npm run dev` in `frontend/` -- starts Vite dev server on `localhost:5173`
2. Run `python main.py` -- detects non-frozen mode, connects QWebEngineView to Vite dev server
3. Hot reload: Vite HMR for frontend; restart Python for backend changes

**Release Build:**
1. `cd frontend && npm run build` -- produces `frontend/dist/`
2. `pyinstaller "Windows Toolbox.spec"` -- produces `dist/Virelo/` with `Virelo.exe`
3. `scripts/build-installer.ps1` -- runs Inno Setup with `installer/virelo.iss`, produces `installer/dist/VireloSetup.exe`

**Dev Mode Detection (`webview.py`):**
- `VIRELO_DEV` env var set to "1"/"true"/"yes" -> dev mode (Vite server)
- `sys.frozen` is False (not PyInstaller) -> dev mode
- `sys.frozen` is True -> release mode (loads `frontend/dist/index.html` via `file://`)

---

*Structure analysis: 2026-04-24*
