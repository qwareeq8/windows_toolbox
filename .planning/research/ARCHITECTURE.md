# Architecture Patterns

**Domain:** PySide6/React/QWebEngine hybrid desktop application
**Researched:** 2026-04-24

## Recommended Architecture

### High-Level Structure

The application follows a **host-shell + embedded-SPA** pattern: a PySide6 (Python) backend acts as the OS integration layer and application host, while a React SPA rendered inside QWebEngineView owns all UI. Communication flows through a single QWebChannel bridge object using JSON payloads.

The refactored architecture preserves this pattern but restructures the Python side from a flat-file monolith into a proper Python package (`virelo/`) with domain-oriented subpackages.

```
+------------------------------------------------------------------+
|  Windows OS                                                      |
|  (Win32, COM, Keyboard hooks, Registry, DWM, Shell)              |
+------------------------------------------------------------------+
        |                           |                    |
+-------v--------+   +-------------v------+   +---------v---------+
| virelo.platform |   | virelo.services    |   | virelo.workers    |
| (Win32 helpers)  |   | (Snap, Explorer,  |   | (QThread workers) |
|                  |   |  Theme, Startup)   |   |                   |
+-------+--------+   +-------+------------+   +---------+---------+
        |                     |                          |
        +----------+----------+----------+---------------+
                   |                     |
          +--------v--------+   +--------v--------+
          | virelo.bridge   |   | virelo.settings |
          | (VireloBridge)  |   | (Settings,      |
          |                 |   |  SettingsState)  |
          +--------+--------+   +-----------------+
                   |
          +--------v--------+
          | QWebChannel     |
          | (JSON payloads) |
          +--------+--------+
                   |
          +--------v--------+
          | React SPA       |
          | (frontend/src/) |
          +--------+--------+
                   |
          +--------v--------+
          | QWebEngineView  |
          | (VireloWebView) |
          +--------+--------+
                   |
          +--------v--------+
          | virelo.app      |
          | (MainWindow,    |
          |  entry point)   |
          +-----------------+
```

### Component Boundaries

| Component | Responsibility | Communicates With |
|-----------|---------------|-------------------|
| `virelo/__main__.py` | Entry point: admin elevation, single-instance mutex, QApplication bootstrap | `virelo.app.MainWindow` |
| `virelo/app.py` | MainWindow: tray icon, window chrome, thread lifecycle orchestration, wiring | Bridge, Services, Workers, WebView |
| `virelo/bridge.py` | QWebChannel QObject: JSON Slots/Signals, input validation, side-effect dispatch | SettingsState, SnapService, MainWindow (for theme/capture callbacks) |
| `virelo/webview.py` | QWebEngineView host: dev/release mode detection, channel setup, JS console routing | Bridge |
| `virelo/settings/` | Settings persistence (QSettings/registry), SettingsState (JSON facade), app_config (defaults/constants) | Bridge, Services |
| `virelo/services/snap.py` | SnapService facade + ShiftSnapRestore engine: keyboard hooks, press timing, Win32 window manipulation | Settings, Platform |
| `virelo/services/explorer.py` | Explorer column autosize: COM IColumnManager, tab-aware state machine | Platform, Workers |
| `virelo/services/theme.py` | Theme resolution: system/dark/light, Windows registry reads, polling | Settings, Bridge |
| `virelo/services/startup.py` | Startup shortcut: WScript.Shell COM, lnk creation/removal | Settings, Platform |
| `virelo/workers/` | QThread worker classes: KeyCaptureWorker, ExplorerAutosizeWorker, ExplorerAutosizeEngine | Services |
| `virelo/platform/` | Pure Win32/ctypes helpers: DPI, monitor rects, DWM, fullscreen detection, window enumeration | Nothing (leaf module) |
| `frontend/src/` | React SPA: settings UI, theme system, command palette | Bridge (via QWebChannel) |

### Data Flow

**Settings Read (startup):**
```
QSettings (registry) --> Settings.__init__() --> SettingsState.get_json()
    --> VireloBridge.get_settings() Slot --> JSON string
    --> QWebChannel --> React bridgeToState() --> useState
```

**Settings Write (user saves):**
```
React state --> stateToBridge() --> JSON string
    --> QWebChannel --> VireloBridge.save_settings() Slot
    --> SettingsState.apply_partial() (validate, coerce, persist)
    --> VireloBridge.settings_changed Signal --> QWebChannel --> React
    --> VireloBridge._apply_side_effects() --> Services
```

**Python-to-React push (theme change, snap status):**
```
Service detects change --> VireloBridge Signal.emit(data)
    --> QWebChannel --> React signal.connect callback --> setState
```

**Data flow direction is strict:** React never calls services directly. The bridge is the only crossing point. Services never import or call the bridge. MainWindow wires services to bridge via signal/slot connections.

## Recommended Python Package Layout

```
virelo/
  __init__.py              # Package marker, version string
  __main__.py              # Entry point (python -m virelo)
  app.py                   # MainWindow: tray, chrome, thread lifecycle, wiring
  bridge.py                # VireloBridge QObject (Slots/Signals)
  webview.py               # VireloWebView, VireloWebPage

  settings/
    __init__.py             # Re-exports Settings, SettingsState, DEFAULTS
    config.py               # APP_NAME, APP_ID, ORGANIZATION, DEFAULTS, LOG constants
    store.py                # Settings class (QSettings read/write)
    state.py                # SettingsState class (JSON facade, validation, apply_partial)

  services/
    __init__.py             # Re-exports service classes
    snap.py                 # SnapService facade + ShiftSnapRestore engine
    explorer.py             # Explorer autosize wrappers (_autosize_quick, _autosize_full)
    theme.py                # resolve_theme, toggle_theme_mode, get_windows_theme, normalize
    startup.py              # create_startup_shortcut, remove_startup_shortcut, path helpers
    capture.py              # CaptureGuard (thread-safe mutex for key capture)

  workers/
    __init__.py             # Re-exports worker classes
    key_capture.py          # KeyCaptureWorker, KeyCaptureSession
    explorer_autosize.py    # ExplorerAutosizeWorker, ExplorerAutosizeEngine, TabAutosizeState

  platform/
    __init__.py             # Re-exports platform helpers
    win32_helpers.py        # USER32/KERNEL32 wrappers, DPI, monitor rects, window rects, DWM
    window_utils.py         # Fullscreen detection, game mode detection, interactive check
    explorer_hwnd.py        # SHELLDLL_DefView/SysListView32 traversal, descendant finders
    com_helpers.py          # _ensure_dispatch, COM cache cleanup
    admin.py                # is_admin, elevation via ShellExecuteW
    resources.py            # resource_path (dev vs PyInstaller path resolution)

  logging.py                # _init_logger, crash log setup, faulthandler

frontend/
  src/
    main.jsx                # React entry: bridge init, theme, Root component
    app.jsx                 # App shell: TitleBar, Sidebar, Footer, page routing, state
    bridge.js               # QWebChannel client: getBridge(), mock bridge
    pages.jsx               # SnapPage, ExplorerPage, ShortcutsPage, GeneralPage, AboutPage
    panels.jsx              # CommandPalette
    primitives.jsx          # Toggle, Button, Card, Row, Segmented, Stepper, Slider, Kbd, Badge
    theme.jsx               # ThemeProvider, useTokens, color palettes, density
    icons.jsx               # 16 inline SVG icons
```

### Layout Rationale

**Why `virelo/` package, not `src/virelo/`:** This is a desktop application built with PyInstaller, not a published library. The flat package layout (`virelo/` at repo root) is simpler for PyInstaller bundling and avoids an unnecessary `src/` indirection layer. The `src/` layout solves import-shadowing problems relevant to libraries on PyPI; that concern does not apply here.

**Why `settings/` subpackage:** The settings pipeline spans three files (`config.py`, `store.py`, `state.py`) with tight coupling. Grouping them under `settings/` with a clean `__init__.py` re-export makes the import story simple: `from virelo.settings import Settings, SettingsState, DEFAULTS`.

**Why `services/` subpackage:** Each service module encapsulates a distinct OS-integration domain (snap, explorer, theme, startup). They share no state with each other. Grouping under `services/` establishes a clear "business logic lives here" boundary.

**Why `platform/` subpackage:** The Win32/ctypes/COM helper functions are pure utility code with no Qt or application state dependency. Isolating them makes them independently testable and prevents circular imports (platform modules never import from services or bridge).

**Why `workers/` subpackage:** Background workers are QThread-aware but logically separate from the services they support. The ExplorerAutosizeEngine (400+ lines) and KeyCaptureSession deserve their own files.

## Recommended React Frontend Layout

The current frontend structure is already well-organized for a single-feature SPA. Keep the flat `frontend/src/` layout -- the app has 7 source files and does not benefit from subdirectories. When the app grows beyond ~15 components, consider splitting `pages.jsx` into individual page files.

```
frontend/src/
  main.jsx           # Entry: bridge init, theme resolution, Root mount
  app.jsx            # Shell (TitleBar, Sidebar, Footer) + state + page routing
  bridge.js          # QWebChannel client (getBridge, useBridgeSync, mock)
  pages.jsx          # All 5 page components (or split later)
  panels.jsx         # CommandPalette
  primitives.jsx     # Shared UI components
  theme.jsx          # ThemeProvider, tokens, palettes
  icons.jsx          # SVG icon registry
```

**No changes needed for the frontend during this refactoring milestone.** The React side is already modular. The bridge client (`bridge.js`) does not need to change because the QWebChannel API surface (`channel.objects.bridge`) is identical regardless of Python package layout.

## Patterns to Follow

### Pattern 1: Service Facade

**What:** Each domain concern gets a service class that owns its state and exposes a narrow public API. The bridge calls service methods; services never import the bridge.

**When:** Any time backend logic needs to be invoked from the UI.

**Why:** The current codebase has the bridge calling `self._main_window._start_key_capture()` and `self._main_window._apply_theme_mode()` -- meaning the bridge depends on MainWindow internals. Services break this coupling.

**Example (after refactoring):**

```python
# virelo/services/snap.py
class SnapService:
    """Facade over ShiftSnapRestore for bridge consumption."""

    def __init__(self, settings: Settings):
        self._settings = settings
        self._engine: Optional[ShiftSnapRestore] = None

    def initialize(self):
        """Create the snap engine. Called after app bootstrap."""
        self._engine = ShiftSnapRestore(self._settings)

    def test_snap(self) -> dict:
        if not self._engine:
            return {"ok": False, "error": "Not initialized"}
        try:
            self._engine.perform(restore=False)
            return {"ok": True, "message": "Snap applied."}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def update_binding(self, key: str): ...
    def update_restore_key(self, key: str): ...
    def cleanup(self): ...
```

```python
# virelo/bridge.py -- bridge calls service, not MainWindow
@Slot(result=str)
def test_snap(self) -> str:
    result = self._snap_service.test_snap()
    self.snap_status.emit(result.get("message", ""), 2000)
    return json.dumps(result)
```

### Pattern 2: MainWindow as Wiring Hub (Not Logic Owner)

**What:** MainWindow creates services, workers, bridge, and webview, then wires them together via signal/slot connections. It does not contain business logic.

**When:** Always. This is the core architectural principle.

**Why:** The current MainWindow is ~485 lines because it owns key capture logic, explorer thread management, theme sync, startup shortcut toggling, and snap test execution. Those belong in services.

**Example (after refactoring):**

```python
# virelo/app.py
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Create
        self._settings = Settings()
        self._settings_state = SettingsState(self._settings)
        self._snap_service = SnapService(self._settings)
        self._theme_service = ThemeService(self._settings)
        self._capture_guard = CaptureGuard()
        self._bridge = VireloBridge(self._settings_state, self._snap_service)

        # Wire
        self._theme_service.theme_changed.connect(self._bridge.theme_applied.emit)
        self._bridge.settings_applied.connect(self._on_settings_applied)

        # Start
        self._snap_service.initialize()
        self._theme_service.start(self._settings.theme)
```

### Pattern 3: Worker Lifecycle via Service

**What:** Services own the lifecycle of their background workers. MainWindow tells a service to start/stop; the service manages its QThread internally.

**When:** For Explorer autosize and key capture.

**Why:** The current MainWindow directly creates QThread, moves workers, connects signals, and stops threads (lines 1098-1151 and 1205-1276). This thread lifecycle code should live in the services that own the workers.

**Example:**

```python
# virelo/services/explorer.py
class ExplorerService(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread: Optional[QThread] = None
        self._worker: Optional[ExplorerAutosizeWorker] = None

    def set_enabled(self, enabled: bool):
        if enabled and not self._is_running():
            self._start_worker()
        elif not enabled and self._is_running():
            self._stop_worker()

    def _start_worker(self):
        self._thread = QThread(self)
        self._worker = ExplorerAutosizeWorker(...)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        # ... signal wiring ...
        self._thread.start()

    def _stop_worker(self):
        if self._worker:
            self._worker.stop()
        if self._thread:
            self._thread.quit()
            self._thread.wait(3000)
        self._worker = None
        self._thread = None
```

### Pattern 4: Single Bridge Object with Domain Grouping

**What:** Keep one `VireloBridge` QObject registered on the QWebChannel. Group Slots logically by comments, not by splitting into multiple registered objects.

**When:** Always for this application.

**Why:** QWebChannel requires objects to be registered before any client connects. Multiple bridge objects add complexity (multiple `channel.objects.*` on the JS side) without benefit for an app with 10-12 Slot methods. The single-bridge pattern is simpler and matches Qt's recommended approach for small-to-medium apps. The bridge already groups its Slots by section comments -- that is sufficient.

### Pattern 5: JSON Envelope for Bridge Communication

**What:** All bridge Slot return values and Signal payloads use a consistent JSON envelope: `{"ok": true, ...data}` or `{"ok": false, "error": "message"}`.

**When:** For every Slot that performs a mutation or action.

**Why:** The current bridge already does this. Preserve the pattern. It allows the React side to handle errors uniformly.

**Example:**

```python
# Python Slot
@Slot(str, result=str)
def save_settings(self, json_str: str) -> str:
    try:
        data = json.loads(json_str)
        result = self._state.apply_partial(data)
        if result["ok"]:
            self.settings_changed.emit(self._state.get_json())
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"ok": False, "error": str(e)})
```

```javascript
// React consumer
bridge.save_settings(JSON.stringify(payload), (response) => {
  const result = JSON.parse(response);
  if (result.ok) { /* update state */ }
  else { /* show error */ }
});
```

## Anti-Patterns to Avoid

### Anti-Pattern 1: Bridge Reaching Into MainWindow Internals

**What:** `self._main_window._start_key_capture()`, `self._main_window._apply_theme_mode()`, `self._main_window.action_run_at_startup.setChecked()` -- the bridge directly calls private methods on MainWindow.

**Why bad:** Creates tight bidirectional coupling between bridge and MainWindow. Breaks testability (cannot test bridge without a MainWindow). Makes refactoring MainWindow dangerous because bridge depends on internal method names.

**Instead:** Bridge calls service methods. MainWindow connects service signals to bridge signals at wiring time. Bridge never holds a reference to MainWindow.

### Anti-Pattern 2: God Module (Current main.py)

**What:** 1450 lines containing entry point, logger setup, Win32 constants, ctypes FFI declarations, 15+ utility functions, ShiftSnapRestore engine (287 lines), startup shortcut logic, MainWindow (485 lines), and the `main()` bootstrap.

**Why bad:** Cannot test any piece in isolation. Cannot reason about dependencies. Cannot reuse utility functions without importing everything. Merge conflicts on every change.

**Instead:** Extract into the package layout described above. Each file under 300 lines. Each module has a clear, single responsibility.

### Anti-Pattern 3: Circular Service References

**What:** Service A imports Service B which imports Service A.

**Why bad:** Python will fail with ImportError. Even if it works via late imports, it signals confused boundaries.

**Instead:** Services depend downward on `platform/` and `settings/`. Services never import each other. If two services need to coordinate, MainWindow wires them via signals.

### Anti-Pattern 4: Worker Creating Its Own Dependencies

**What:** A QThread worker importing and instantiating COM objects, platform helpers, or services inside its `run()` method.

**Why bad:** Makes the worker hard to test and tightly coupled to the platform layer.

**Instead:** Inject callable dependencies into the worker's constructor. The current `ExplorerAutosizeWorker` already does this correctly (accepts `autosize_quick_fn`, `autosize_full_fn`, `is_interactive_fn` callables). Preserve this pattern.

## Dependency Graph (Import Direction)

```
virelo/__main__.py
  |
  v
virelo/app.py (MainWindow)
  |
  +---> virelo/bridge.py (VireloBridge)
  |       |
  |       +---> virelo/settings/ (SettingsState, config)
  |       +---> virelo/services/snap.py (SnapService)
  |
  +---> virelo/webview.py (VireloWebView)
  |       |
  |       +---> virelo/bridge.py
  |
  +---> virelo/services/snap.py
  |       |
  |       +---> virelo/settings/ (Settings)
  |       +---> virelo/platform/ (Win32 helpers)
  |
  +---> virelo/services/explorer.py
  |       |
  |       +---> virelo/platform/ (COM, window utils)
  |       +---> virelo/workers/ (ExplorerAutosizeWorker)
  |
  +---> virelo/services/theme.py
  |       |
  |       +---> virelo/settings/ (config)
  |       +---> virelo/platform/ (registry read)
  |
  +---> virelo/services/startup.py
  |       |
  |       +---> virelo/platform/ (COM, resources)
  |
  +---> virelo/settings/ (Settings, SettingsState)
  +---> virelo/logging.py

virelo/platform/*  --> (no virelo imports -- leaf modules only)
virelo/workers/*   --> (no virelo imports except platform callables injected at construction)
```

**Key rule:** Arrows point downward only. No module imports from a module above it in the graph. `platform/` is the bottom layer, imported by everything but importing nothing from `virelo/`.

## Build Order Implications (Refactoring Sequence)

The refactoring must proceed bottom-up to avoid breaking the running application at any step.

### Phase 1: Create package skeleton + move leaf modules

1. Create `virelo/` directory with `__init__.py` and `__main__.py`
2. Move `platform/` helpers first (pure functions, no app state, no Qt dependency except types)
   - Extract Win32 constants, ctypes declarations, and utility functions from `main.py` lines 165-530
   - Move existing `theme.py` functions into `services/theme.py`
   - Move `capture_guard.py` to `services/capture.py`
   - Move `startup_shortcut.py` to `services/startup.py`
3. Move `settings/` (config, store, state) -- these are already separate files, just need directory restructuring
4. Update imports in all consumers

**Why first:** Platform and settings have zero dependencies on other `virelo/` modules. Moving them cannot create circular imports.

### Phase 2: Move services

5. Move `ShiftSnapRestore` from `main.py` into `services/snap.py`, merge with existing `snap_service.py`
6. Extract explorer autosize wrappers from `main.py` into `services/explorer.py`
7. Extract startup shortcut functions from `main.py` into `services/startup.py`
8. Move `workers.py` contents into `workers/` subpackage (split into `key_capture.py` and `explorer_autosize.py`)

**Why second:** Services depend on platform and settings (already moved). Workers depend on platform helpers (already moved).

### Phase 3: Restructure bridge and app

9. Move `bridge.py` into `virelo/bridge.py` -- update imports to use new service locations
10. Eliminate `bridge._main_window` dependency: move key capture orchestration to a `CaptureService`, theme calls to `ThemeService`, startup calls to `StartupService`
11. Move `webview.py` into `virelo/webview.py`
12. Move MainWindow from `main.py` into `virelo/app.py` -- strip all extracted logic, keep only wiring

**Why last:** Bridge and MainWindow depend on services, settings, and platform. They must move after their dependencies are in place.

### Phase 4: Entry point and cleanup

13. Move `main()` function into `virelo/__main__.py`
14. Delete the old root-level `.py` files (main.py, bridge.py, etc.)
15. Update PyInstaller spec to point to `virelo/__main__.py`
16. Verify dev mode (`python -m virelo`) and frozen mode (PyInstaller) both work

## Scalability Considerations

| Concern | Current (1 user) | Future (module registry) |
|---------|-------------------|--------------------------|
| Bridge size | 12 Slots, manageable | Could register multiple QObjects per module |
| Settings | 11 keys, flat dict | Could namespace per module: `snap.key`, `explorer.auto_size` |
| Workers | 2 background threads | Each module owns its workers |
| Frontend | Single SPA, 5 pages | Could lazy-load module UIs as separate chunks |

The package layout supports the "module registry" vision from PROJECT.md without requiring it now. Each service subpackage could become a self-contained module with its own settings namespace, bridge slots, and workers.

## Sources

- [Qt QWebChannel Documentation](https://doc.qt.io/qt-6/qwebchannel.html) - Object registration, lifecycle, API design (HIGH confidence)
- [KDAB: Qt WebChannel - bridging the gap](https://www.kdab.com/qt-webchannel-bridging-gap-cqml-web/) - Architecture recommendations, async communication model (HIGH confidence)
- [PySide6 QWebChannel Documentation](https://doc.qt.io/qtforpython-6/PySide6/QtWebChannel/QWebChannel.html) - Python-specific API (HIGH confidence)
- [PySide6 QThread Documentation](https://doc.qt.io/qtforpython-6/PySide6/QtCore/QThread.html) - Worker + moveToThread pattern (HIGH confidence)
- [PySide6 Best Practices (ZynU)](https://www.zynu.net/ai-skills/pyside6-best-practices) - MainWindow as wiring hub, signal bus, widget lifecycle (MEDIUM confidence)
- [Python Application Layouts (Real Python)](https://realpython.com/python-application-layouts/) - Package vs flat layout guidance (HIGH confidence)
- [Remote Frontends for PySide-based Tooling](https://www.call-with.cc/post/remote-frontends-for-pyside2-based-vfx-tooling-over) - Bridge as "Core" pattern, JSON dispatch (MEDIUM confidence)
- [pywebchannel](https://pywebchannel.readthedocs.io/) - TypeScript generation for QWebChannel bridges (MEDIUM confidence)
- [src layout vs flat layout (Python Packaging User Guide)](https://packaging.python.org/en/latest/discussions/src-layout-vs-flat-layout/) - Layout trade-offs for PyInstaller apps (HIGH confidence)

---

*Architecture analysis: 2026-04-24*
