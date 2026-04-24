# Phase 3: Structure and Quality - Pattern Map

**Mapped:** 2026-04-24
**Files analyzed:** 42 new/modified files
**Analogs found:** 36 / 42

## File Classification

### Package init files (no analog needed -- mechanical)

| New File | Role | Data Flow | Closest Analog | Match Quality |
|----------|------|-----------|----------------|---------------|
| `virelo/__init__.py` | config | -- | -- | boilerplate |
| `virelo/app/__init__.py` | config | -- | -- | boilerplate |
| `virelo/bridge/__init__.py` | config | -- | -- | boilerplate |
| `virelo/services/__init__.py` | config | -- | -- | boilerplate |
| `virelo/workers/__init__.py` | config | -- | -- | boilerplate |
| `virelo/platform/__init__.py` | config | -- | -- | boilerplate |
| `virelo/settings/__init__.py` | config | -- | -- | boilerplate |
| `tests/__init__.py` | config | -- | -- | boilerplate |
| `tests/unit/__init__.py` | config | -- | -- | boilerplate |
| `tests/integration/__init__.py` | config | -- | -- | boilerplate |

### Core Python restructuring

| New/Modified File | Role | Data Flow | Closest Analog | Match Quality |
|-------------------|------|-----------|----------------|---------------|
| `virelo/app/config.py` | config | -- | `app_config.py` (lines 1-37) | exact-move |
| `virelo/app/__main__.py` | utility | request-response | `main.py` (lines 1-138, 1345-1449) | extract |
| `virelo/app/window.py` | controller | event-driven | `main.py` (lines 905-1019) | extract |
| `virelo/bridge/bridge.py` | controller | request-response | `bridge.py` (lines 1-282) | exact-move |
| `virelo/bridge/capture_guard.py` | utility | event-driven | `capture_guard.py` (lines 1-17) | exact-move |
| `virelo/services/snap.py` | service | request-response | `snap_service.py` (lines 1-49) + `main.py` (lines 609-700) | merge |
| `virelo/services/explorer_columns.py` | service | CRUD | `explorer_columns.py` (all) | exact-move |
| `virelo/platform/theme.py` | service | request-response | `theme.py` (lines 1-38) | exact-move |
| `virelo/platform/startup.py` | utility | file-I/O | `startup_shortcut.py` (lines 1-17) | exact-move |
| `virelo/workers/key_capture.py` | worker | event-driven | `workers.py` (lines 78-141, 636-664) | extract |
| `virelo/workers/explorer.py` | worker | event-driven | `workers.py` (lines 1-77, 143-978) | extract |
| `virelo/platform/win32_helpers.py` | utility | request-response | `main.py` (lines 165-528) | extract |
| `virelo/platform/resources.py` | utility | file-I/O | `main.py` (lines 198-201) + `webview.py` (lines 54-67) | consolidate |
| `virelo/platform/paths.py` | utility | transform | `workers.py` (lines 143-171) + `explorer_columns.py` (line 850) | consolidate |
| `virelo/settings/persistence.py` | model | CRUD | `settings.py` (lines 1-101) | exact-move |
| `virelo/settings/state.py` | model | CRUD | `settings_state.py` (lines 1-136) | exact-move |
| `main.py` (modified) | utility | -- | `main.py` current (lines 198-201 shim pattern) | rewrite-to-shim |
| `Virelo.spec` (modified) | config | -- | `Virelo.spec` current (all) | modify |

### Test infrastructure

| New File | Role | Data Flow | Closest Analog | Match Quality |
|----------|------|-----------|----------------|---------------|
| `tests/conftest.py` | test | -- | RESEARCH.md Pattern 2 | from-research |
| `tests/unit/test_settings_state.py` | test | CRUD | `settings_state.py` (validation logic) | role-match |
| `tests/unit/test_theme.py` | test | transform | `theme.py` (pure functions) | role-match |
| `tests/unit/test_snap_geometry.py` | test | transform | `main.py` (lines 275-288, 344-351) | role-match |
| `tests/unit/test_app_config.py` | test | transform | `app_config.py` (lines 31-37) | role-match |
| `tests/unit/test_bridge_payload.py` | test | request-response | `bridge.py` (JSON envelope pattern) | role-match |
| `tests/unit/test_paths.py` | test | transform | `workers.py` (lines 143-171) | role-match |
| `tests/unit/test_capture_guard.py` | test | event-driven | `capture_guard.py` (thread safety) | role-match |
| `tests/integration/conftest.py` | test | -- | RESEARCH.md Pitfall 6 | from-research |

### Frontend test infrastructure

| New/Modified File | Role | Data Flow | Closest Analog | Match Quality |
|-------------------|------|-----------|----------------|---------------|
| `frontend/vite.config.js` (modified) | config | -- | `frontend/vite.config.js` current | modify |
| `frontend/src/test-setup.js` | config | -- | RESEARCH.md Vitest config | from-research |
| `frontend/src/__tests__/app.test.jsx` | test | transform | `frontend/src/app.jsx` (lines 148-177) | role-match |
| `frontend/src/__tests__/panels.test.jsx` | test | transform | `frontend/src/panels.jsx` (lines 19-40) | role-match |
| `frontend/src/__tests__/primitives.test.jsx` | test | -- | `frontend/src/primitives.jsx` (components) | role-match |

### Quality tooling and CI

| New File | Role | Data Flow | Closest Analog | Match Quality |
|----------|------|-----------|----------------|---------------|
| `pyproject.toml` | config | -- | RESEARCH.md Code Examples | from-research |
| `.github/workflows/ci.yml` | config | -- | RESEARCH.md Code Examples | from-research |

---

## Pattern Assignments

### `virelo/app/config.py` (config, exact-move)

**Analog:** `app_config.py` (entire file, 37 lines)

This is a direct move. No structural changes. All constants and the `normalize_snap_presses` function transfer as-is.

**Full module pattern** (lines 1-37):
```python
APP_NAME = "Virelo"
ORGANIZATION = "Yusuf Qwareeq"
APP_DISPLAY_NAME = "Virelo"
APP_VERSION = "1.5.0"
APP_EXECUTABLE_NAME = "Virelo.exe"
APP_DIST_DIR_NAME = "Virelo"
APP_PUBLISHER = "Yusuf Qwareeq"
APP_SUPPORT_URL = "https://github.com/yusufqwareeq/virelo"
APP_SETTINGS_ORG = "Yusuf Qwareeq"
APP_LOG_DIR = "Virelo"
APP_LOG_FILE = "virelo.log"
APP_ID = "com.yusufqwareeq.virelo"
LOG_DIR = "Virelo"
LOG_FILE = "virelo.log"
SETTINGS_GROUP = "Settings"

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

def normalize_snap_presses(value):
    try:
        val = int(value)
    except Exception:
        return DEFAULTS["snap_presses"]
    return max(1, val)
```

---

### `virelo/app/__main__.py` (utility, extract from `main.py`)

**Analog:** `main.py` lines 1-138 (imports, logging, crash diagnostics) + lines 1345-1449 (entry point `main()`)

**Imports pattern** (lines 1-56 of `main.py`):
```python
import ctypes
import logging
import os
import sys
import threading
import time
from logging.handlers import RotatingFileHandler

# Exit early on non-Windows platforms.
if sys.platform != "win32":
    print("Virelo requires Windows.")
    sys.exit(1)

import atexit
import faulthandler
```

**Logger init pattern** (lines 62-111 of `main.py`):
```python
def _init_logger() -> logging.Logger:
    """Initialize a rotating file logger in a per-user location."""
    base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    log_dir = os.path.join(base, LOG_DIR)
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, LOG_FILE)
    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    # ... rotating file handler + console handler setup
    return logger
```

**Crash diagnostics pattern** (lines 114-138 of `main.py`):
```python
LOG = _init_logger()

_CRASH_LOG = None
try:
    crash_log_path = os.path.join(
        os.path.dirname(getattr(LOG, "log_path", "")), "crash.log"
    )
    _CRASH_LOG = open(crash_log_path, "a", encoding="utf-8")
    faulthandler.enable(_CRASH_LOG)
except Exception:
    try:
        faulthandler.enable()
    except Exception:
        pass

def _cleanup_faulthandler():
    if _CRASH_LOG:
        try:
            _CRASH_LOG.close()
        except Exception:
            pass

atexit.register(_cleanup_faulthandler)
```

**Note:** Updated imports will reference `virelo.app.config` instead of `app_config`, `virelo.app.window` instead of inline `MainWindow`, etc.

---

### `virelo/app/window.py` (controller, event-driven, extract from `main.py`)

**Analog:** `main.py` lines 905-1019 (class declaration, `__init__`)

**Class declaration and init pattern** (lines 905-916):
```python
class MainWindow(QtWidgets.QMainWindow):
    """Main application window with tray icon and QWebEngineView frontend."""

    key_captured = QtCore.Signal(str)
    snap_key_status = QtCore.Signal(str, int)

    def __init__(self):
        super().__init__()
        self.settings = Settings()
```

**Bridge wiring pattern** (lines 991-1004):
```python
        # --- Bridge + WebView ---
        self._settings_state = SettingsState(self.settings)
        self._snap_service = SnapService(None)  # shift_mgr set after construction
        self._bridge = VireloBridge(self._settings_state, self._snap_service, parent=self)
        self._bridge.set_main_window(self)
        self._bridge.set_capture_guard(self._capture_guard)

        self.webview = VireloWebView(self._bridge, parent=self)
        self.setCentralWidget(self.webview)
        self.snap_key_status.connect(self._bridge.snap_status.emit)
```

**Import update:** All imports must change from flat-module names to package paths:
- `from app_config import ...` becomes `from virelo.app.config import ...`
- `from settings import Settings` becomes `from virelo.settings import Settings`
- `from bridge import VireloBridge` becomes `from virelo.bridge import VireloBridge`
- etc.

---

### `virelo/bridge/bridge.py` (controller, request-response, exact-move)

**Analog:** `bridge.py` (entire file, 282 lines)

**Imports pattern** (lines 1-23):
```python
import json
import logging
from typing import Optional

from PySide6 import QtCore
from PySide6.QtCore import QObject, Signal, Slot

from settings_state import SettingsState
from snap_service import SnapService
```

After move, imports become:
```python
from virelo.settings.state import SettingsState
from virelo.services.snap import SnapService
```

**JSON envelope pattern** (used in every Slot, lines 63-71):
```python
@Slot(result=str)
def get_settings(self) -> str:
    try:
        settings = self._state.get_all()
        return json.dumps({"ok": True, "data": settings})
    except Exception as e:
        LOG.exception("get_settings failed")
        return json.dumps({"ok": False, "error": str(e)})
```

**Input validation pattern** (lines 80-93):
```python
@Slot(str, result=str)
def save_settings(self, json_str: str) -> str:
    try:
        data = json.loads(json_str)
        if not isinstance(data, dict):
            return json.dumps({"ok": False, "error": "Expected JSON object"})
        # ... validate and process
    except json.JSONDecodeError as e:
        return json.dumps({"ok": False, "error": f"Invalid JSON: {e}"})
    except Exception as e:
        LOG.exception("save_settings failed")
        return json.dumps({"ok": False, "error": str(e)})
```

---

### `virelo/services/snap.py` (service, request-response, merge)

**Analog:** `snap_service.py` (lines 1-49) for the facade + `main.py` (lines 609-700) for `ShiftSnapRestore`

**Service facade pattern** (`snap_service.py` lines 10-49):
```python
LOG = logging.getLogger("Virelo")

class SnapService:
    """Narrow API surface for snap/restore actions."""

    def __init__(self, shift_mgr):
        self._mgr = shift_mgr

    def set_manager(self, mgr):
        self._mgr = mgr

    def test_snap(self) -> dict:
        if self._mgr is None:
            return {"ok": False, "error": "Snap manager not initialized"}
        try:
            self._mgr.perform(False)
            return {"ok": True, "message": "Snap test applied to the active window."}
        except Exception as e:
            LOG.exception("test_snap failed")
            return {"ok": False, "error": str(e)}
```

**ShiftSnapRestore pattern** (`main.py` lines 609-630):
```python
class ShiftSnapRestore(QtCore.QObject):
    triggered = QtCore.Signal(bool)
    blocked = QtCore.Signal(str)

    def __init__(self, settings: Settings):
        super().__init__()
        self.settings = settings
        self._press_times: Deque[float] = deque(
            maxlen=normalize_snap_presses(self.settings.snap_presses)
        )
        self._press_lock = threading.Lock()
        self._held = False
        self._orig_sizes: Dict[int, Dict[str, Union[Tuple[int, int, int, int], bool]]] = {}
        self.current_key = str(settings.snap_key)
        self.restore_key = str(getattr(settings, "restore_key", "ctrl"))
        self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
        self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)
        self._fetch_open_windows()
```

**New testable extraction** (`calculate_snap_position` from RESEARCH.md Pattern 3):
```python
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

---

### `virelo/platform/theme.py` (service, exact-move)

**Analog:** `theme.py` (entire file, 38 lines)

**Dependency injection pattern** (lines 23-37):
```python
def get_windows_theme(read_registry=None):
    try:
        if read_registry is None:
            import winreg
            # ... actual registry read
        else:
            value = read_registry()
        return "light" if int(value) == 1 else "dark"
    except Exception:
        return "dark"
```

This injectable `read_registry` parameter is the project's established testability pattern. Other modules should follow this approach.

---

### `virelo/settings/persistence.py` (model, CRUD, exact-move)

**Analog:** `settings.py` (entire file, 101 lines)

**Defensive coercion pattern** (lines 79-101):
```python
def _safe_int(val, default):
    try:
        if val is None:
            return default
        return int(val)
    except Exception:
        return default

def _safe_bool(val, default):
    if isinstance(val, bool):
        return val
    if val is None:
        return default
    text = str(val).strip().lower()
    if text in ("1", "true", "yes", "on"):
        return True
    if text in ("0", "false", "no", "off"):
        return False
    try:
        return bool(int(val))
    except Exception:
        return default
```

---

### `virelo/settings/state.py` (model, CRUD, exact-move)

**Analog:** `settings_state.py` (entire file, 136 lines)

**KEYS validation dict pattern** (lines 30-42):
```python
KEYS = {
    "snap_key":          (str,  None),
    "restore_key":       (str,  None),
    "enable_snap":       (bool, None),
    "snap_presses":      (int,  (1, 10)),
    "snap_interval":     (int,  (100, 5000)),
    "width_pct":         (int,  (10, 100)),
    "height_pct":        (int,  (10, 100)),
    "ex_auto_size":      (bool, None),
    "game_mode_enabled": (bool, None),
    "run_at_startup":    (bool, None),
    "theme":             (str,  None),
}
```

**Draft/commit model pattern** (lines 79-128):
```python
def apply_draft(self, data: dict) -> dict:
    unknown = [k for k in data if k not in self.KEYS]
    if unknown:
        return {"ok": False, "error": f"Unknown keys: {unknown}"}
    validated = {}
    for key, value in data.items():
        coercer, bounds = self.KEYS[key]
        try:
            coerced = coercer(value)
        except (ValueError, TypeError) as e:
            return {"ok": False, "error": f"Invalid type for {key}: {e}"}
        if bounds is not None:
            lo, hi = bounds
            if not (lo <= coerced <= hi):
                return {"ok": False, "error": f"{key} must be between {lo} and {hi}, got {coerced}"}
        validated[key] = coerced
    # ... store in draft
    return {"ok": True, "applied": validated}
```

---

### `virelo/platform/win32_helpers.py` (utility, extract from `main.py`)

**Analog:** `main.py` lines 165-528

**Win32 constants pattern** (lines 165-191):
```python
USER32 = ctypes.windll.user32
KERNEL32 = ctypes.windll.kernel32

LVM_FIRST = 0x1000
LVM_GETHEADER = LVM_FIRST + 31
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_SETCOLUMNWIDTH = LVM_FIRST + 30

LVSCW_AUTOSIZE = -1
LVSCW_AUTOSIZE_USEHEADER = -2
# ... more constants
```

**Pure geometry function pattern** (lines 275-288):
```python
FULLSCREEN_TOLERANCE = 3

def _rect_matches_monitor(
    rect: Tuple[int, int, int, int], monitor: Tuple[int, int, int, int]
) -> bool:
    left, top, right, bottom = rect
    left_edge, top_edge, right_edge, bottom_edge = monitor
    return (
        abs(left - left_edge) <= FULLSCREEN_TOLERANCE
        and abs(top - top_edge) <= FULLSCREEN_TOLERANCE
        and abs(right - right_edge) <= FULLSCREEN_TOLERANCE
        and abs(bottom - bottom_edge) <= FULLSCREEN_TOLERANCE
    )
```

**Game detection pattern** (lines 334-351):
```python
def _looks_like_game_window(hwnd: int) -> bool:
    try:
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    except Exception:
        return False
    has_caption = bool(style & win32con.WS_CAPTION or style & win32con.WS_BORDER)
    is_popup = bool(style & win32con.WS_POPUP)
    return is_popup and not has_caption

def _should_skip_snap_for_game(
    hwnd: int, settings, full_screen: bool
) -> bool:
    return (
        full_screen
        and getattr(settings, "game_mode_enabled", True)
        and _looks_like_game_window(hwnd)
    )
```

---

### `virelo/platform/resources.py` (utility, consolidate duplicates)

**Analog:** `main.py` line 198-201 + `webview.py` lines 54-67 (two copies of `resource_path`)

**Pattern to consolidate** (`main.py` lines 198-201):
```python
def resource_path(relative_path: str) -> str:
    """Return absolute path to resource, works for dev and PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)
```

**Second copy** (`webview.py` lines 54-67 -- same logic, different docstring). Both copies converge into one function in `virelo/platform/resources.py`. The `__file__` reference must be updated to point to the project root, not the module's own directory.

---

### `virelo/platform/paths.py` (utility, consolidate duplicates)

**Analog:** `workers.py` lines 143-171 + `explorer_columns.py` line 850 (two copies of `canonicalize_path`)

**Pattern to consolidate** (`workers.py` lines 143-171):
```python
def _canonicalize_path(path: str) -> str:
    if not path:
        return ""
    path = path.strip()
    # Handle file:/// URLs
    if path.lower().startswith("file:///"):
        path = urllib.parse.unquote(path[8:])
    # Normalize separators
    path = path.replace("/", "\\")
    # Strip trailing separators (but keep root like C:\)
    while len(path) > 3 and path.endswith("\\"):
        path = path[:-1]
    # Normalize case
    path = path.lower()
    return path
```

Both `workers.py._canonicalize_path` and `explorer_columns.canonicalize_path` converge into a single `canonicalize_path()` in this module.

---

### `virelo/workers/key_capture.py` (worker, event-driven, extract from `workers.py`)

**Analog:** `workers.py` lines 78-141 (KeyCaptureSession) + lines 636-664 (KeyCaptureWorker)

**KeyCaptureSession pure-logic pattern** (lines 78-99):
```python
class KeyCaptureSession:
    def __init__(self, keyboard, cancel_key="esc", timeout_s=15.0, poll_interval=0.01):
        self._keyboard = keyboard
        self._cancel_key = cancel_key
        self._timeout_s = timeout_s
        self._poll_interval = poll_interval
        self._done = threading.Event()
        self._stop = threading.Event()
        self._result = None
        self._reason = None
```

**QObject worker pattern** (lines 636-664):
```python
try:
    from PySide6 import QtCore
except Exception:
    QtCore = None

if QtCore is not None:
    class KeyCaptureWorker(QtCore.QObject):
        captured = QtCore.Signal(str)
        cancelled = QtCore.Signal(str)
        finished = QtCore.Signal()

        def __init__(self, keyboard_module=None, cancel_key="esc", timeout_s=15.0):
            super().__init__()
            if keyboard_module is None:
                import keyboard as keyboard_module
            self._session = KeyCaptureSession(
                keyboard_module, cancel_key=cancel_key, timeout_s=timeout_s,
            )

        @QtCore.Slot()
        def run(self):
            key, reason = self._session.run()
            if key:
                self.captured.emit(key)
            else:
                self.cancelled.emit(reason)
            self.finished.emit()
```

**Key pattern:** The conditional PySide6 import (`try/except` guard) enables unit tests to test `KeyCaptureSession` without PySide6 installed. This pattern must be preserved.

---

### `virelo/workers/explorer.py` (worker, event-driven, extract from `workers.py`)

**Analog:** `workers.py` lines 1-77 (constants, dataclasses), 143-635 (engine), 666-978 (ExplorerAutosizeWorker)

**Dataclass state pattern** (lines 49-76):
```python
@dataclass
class TabAutosizeState:
    tab_id: int
    hwnd: int
    path: str = ""
    view_mode: Optional[int] = None
    navigation_token: int = 0
    first_seen_at: float = 0.0
    path_stable_since: float = 0.0
    last_autosize_attempt: float = 0.0
    # ... more state fields
    consecutive_failures: int = 0
    circuit_open_until: float = 0.0
```

**COM lifecycle pattern** (`ExplorerAutosizeWorker.run()`, lines 710-735):
```python
@QtCore.Slot()
def run(self):
    import pythoncom
    import win32com.client
    # COM is initialized ONCE for this worker thread (STA)
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:
        pythoncom.CoInitialize()
    # Cache Shell.Application to avoid creating/destroying it every poll cycle
    shell_app = None
    try:
        shell_app = win32com.client.Dispatch("Shell.Application")
    except Exception as e:
        log.error("Explorer autosize worker: failed to create Shell.Application: %s", e)
```

**Critical constraint (D-07):** COM init, Shell.Application caching, `iter_tabs`, and all COM-dependent closures must remain co-located in this single file. Do not separate COM initialization from COM usage across modules.

---

### `main.py` (modified -- rewrite to thin shim)

**Analog:** Current `main.py` entry point pattern

**Target shim** (per D-03):
```python
from virelo.app import main

main()
```

---

### `Virelo.spec` (modified)

**Analog:** Current `Virelo.spec` (lines 1-60)

**Version parsing pattern** (lines 6-10 -- must update path):
```python
# Parse APP_VERSION via regex -- do NOT import app_config directly.
_cfg = Path("app_config.py").read_text()
_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', _cfg)
```

After restructuring, this becomes:
```python
_cfg = Path("virelo/app/config.py").read_text()
_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', _cfg)
```

**Hidden imports update** (lines 20-23):
```python
hiddenimports=[
    'virelo', 'virelo.app', 'virelo.app.window', 'virelo.app.config',
    'virelo.bridge', 'virelo.bridge.bridge', 'virelo.bridge.capture_guard',
    'virelo.services', 'virelo.services.snap', 'virelo.services.theme',
    'virelo.settings', 'virelo.settings.persistence', 'virelo.settings.state',
    'virelo.workers', 'virelo.workers.key_capture', 'virelo.workers.explorer',
    'virelo.platform', 'virelo.platform.win32_helpers', 'virelo.platform.resources',
    'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineCore', 'PySide6.QtWebChannel',
],
```

---

### `virelo/__init__.py` (package marker)

**Pattern:** Re-export `__version__` from config:
```python
from virelo.app.config import APP_VERSION as __version__
```

### Subpackage `__init__.py` files (re-export pattern)

**Ruff F401 note:** All `__init__.py` re-exports will trigger `F401` (imported but unused). The `pyproject.toml` must include `per-file-ignores` to suppress this:
```toml
[tool.ruff.lint.per-file-ignores]
"virelo/*/__init__.py" = ["F401"]
```

**Example re-export pattern** (`virelo/app/__init__.py`):
```python
from virelo.app.config import APP_NAME, APP_VERSION, DEFAULTS
from virelo.app.window import MainWindow
```

**Example re-export pattern** (`virelo/settings/__init__.py`):
```python
from virelo.settings.persistence import Settings
from virelo.settings.state import SettingsState
```

---

### `tests/conftest.py` (test fixtures)

**Analog:** `settings_state.py` (MockSettings derives from its interface) + RESEARCH.md Pattern 2

**MockSettings pattern** (from RESEARCH.md, verified against `settings.py` interface):
```python
import pytest
from virelo.app.config import DEFAULTS

class MockSettings:
    """In-memory Settings replacement for unit tests."""
    def __init__(self, **overrides):
        for key, val in DEFAULTS.items():
            setattr(self, key, overrides.get(key, val))
    def save(self):
        pass
    def clear(self):
        pass

@pytest.fixture
def mock_settings():
    return MockSettings()

@pytest.fixture
def settings_state(mock_settings):
    from virelo.settings.state import SettingsState
    return SettingsState(mock_settings)
```

---

### `tests/unit/test_settings_state.py` (test, CRUD)

**Analog:** `settings_state.py` validation logic (lines 79-112)

**Test pattern:** Test the `apply_draft` / `commit_draft` / `discard_draft` cycle against `MockSettings`. Tests exercise the `KEYS` validation dict, range bounds, type coercion, unknown key rejection, and the `{ok, data/error}` envelope.

```python
def test_apply_draft_validates_range(state):
    result = state.apply_draft({"width_pct": 150})
    assert result["ok"] is False
    assert "must be between" in result["error"]

def test_apply_draft_rejects_unknown_keys(state):
    result = state.apply_draft({"nonexistent_key": "value"})
    assert result["ok"] is False

def test_apply_draft_coerces_types(state):
    result = state.apply_draft({"snap_presses": "5"})
    assert result["ok"] is True
    assert result["applied"]["snap_presses"] == 5
```

---

### `tests/unit/test_theme.py` (test, transform)

**Analog:** `theme.py` pure functions (lines 1-21)

**Test pattern:** All functions in `theme.py` are pure (no side effects except `get_windows_theme` which uses DI). Tests call them directly.

```python
from virelo.services.theme import normalize_theme_mode, resolve_theme, toggle_theme_mode

def test_normalize_valid_modes():
    assert normalize_theme_mode("dark") == "dark"
    assert normalize_theme_mode("light") == "light"
    assert normalize_theme_mode("system") == "system"

def test_normalize_invalid_falls_back():
    assert normalize_theme_mode("invalid") == "system"
```

---

### `tests/unit/test_snap_geometry.py` (test, transform)

**Analog:** `main.py` lines 275-288 (`_rect_matches_monitor`) + lines 344-351 (`_should_skip_snap_for_game`)

**Test pattern:** Pure math functions tested with boundary values.

```python
def test_rect_matches_monitor_exact():
    from virelo.platform.win32_helpers import _rect_matches_monitor
    assert _rect_matches_monitor((0, 0, 1920, 1080), (0, 0, 1920, 1080)) is True

def test_rect_matches_monitor_within_tolerance():
    from virelo.platform.win32_helpers import _rect_matches_monitor
    assert _rect_matches_monitor((-2, -1, 1922, 1081), (0, 0, 1920, 1080)) is True

def test_rect_does_not_match_monitor():
    from virelo.platform.win32_helpers import _rect_matches_monitor
    assert _rect_matches_monitor((100, 100, 800, 600), (0, 0, 1920, 1080)) is False
```

---

### `tests/unit/test_capture_guard.py` (test, event-driven)

**Analog:** `capture_guard.py` (lines 1-17)

**Test pattern:** Thread-safety testing with `CaptureGuard`.

```python
from virelo.bridge.capture_guard import CaptureGuard

def test_try_start_succeeds_first_time():
    guard = CaptureGuard()
    assert guard.try_start() is True

def test_try_start_fails_while_active():
    guard = CaptureGuard()
    guard.try_start()
    assert guard.try_start() is False

def test_finish_allows_restart():
    guard = CaptureGuard()
    guard.try_start()
    guard.finish()
    assert guard.try_start() is True
```

---

### `frontend/vite.config.js` (modified -- add test config)

**Analog:** Current `frontend/vite.config.js` (lines 1-23)

**Addition pattern** (add `test` key to existing config):
```javascript
export default defineConfig({
  // ... existing plugins, base, define, build, server ...
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: './src/test-setup.js',
  },
});
```

---

### `frontend/src/__tests__/app.test.jsx` (test, transform)

**Analog:** `frontend/src/app.jsx` lines 148-177 (`bridgeToState`, `stateToBridge`)

**Note:** These functions are currently not exported from `app.jsx`. They must be exported as named exports for testing (A1 assumption from RESEARCH.md). They are pure mapping functions with no side effects.

**Test pattern:**
```javascript
import { describe, it, expect } from 'vitest';

describe('bridgeToState', () => {
  it('maps Python keys to React keys', async () => {
    const { bridgeToState } = await import('../app.jsx');
    const result = bridgeToState({
      enable_snap: true,
      snap_key: 'shift',
      // ... all keys
    });
    expect(result.snapEnabled).toBe(true);
    expect(result.snapKey).toBe('SHIFT');
  });
});
```

---

### `frontend/src/__tests__/panels.test.jsx` (test, transform)

**Analog:** `frontend/src/panels.jsx` lines 19-40 (commands array + filter logic)

**Test target:** The filtering logic `commands.filter(c => c.label.toLowerCase().includes(q.toLowerCase()))` is embedded in the component. Testing requires rendering `CommandPalette` with a mock `app` prop and verifying filtered output.

**Pattern:** Use `@testing-library/react` to render and query:
```javascript
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
```

---

### `frontend/src/__tests__/primitives.test.jsx` (test, smoke)

**Analog:** `frontend/src/primitives.jsx` (component definitions)

**Test pattern:** Render smoke tests verifying components mount without error. Requires wrapping in a theme provider since primitives use `useTokens()`.

---

## Shared Patterns

### Logging

**Source:** `main.py` line 114, `bridge.py` line 24, `snap_service.py` line 10
**Apply to:** All Python modules that log

```python
import logging
LOG = logging.getLogger("Virelo")
```

Every module uses `logging.getLogger("Virelo")` -- the logger name is always `"Virelo"`, matching `APP_NAME`.

### JSON Envelope

**Source:** `bridge.py` lines 63-71, `settings_state.py` lines 79-112, `snap_service.py` lines 24-32
**Apply to:** All bridge/service return values

Success: `{"ok": True, "data": ...}` or `{"ok": True, "applied": ...}`
Failure: `{"ok": False, "error": "..."}`

### Error Handling

**Source:** `bridge.py` (every Slot method), `snap_service.py` lines 26-32
**Apply to:** All controller/service methods

```python
try:
    # business logic
    return json.dumps({"ok": True, "data": result})
except Exception as e:
    LOG.exception("method_name failed")
    return json.dumps({"ok": False, "error": str(e)})
```

### Dependency Injection for Testability

**Source:** `theme.py` line 23 (`read_registry=None`), `startup_shortcut.py` line 4 (`exists=os.path.exists`)
**Apply to:** Any function that touches platform APIs and needs unit testing

```python
def get_windows_theme(read_registry=None):
    try:
        if read_registry is None:
            # real implementation
        else:
            value = read_registry()
```

### Import Update Map

**Apply to:** All files being moved into the `virelo/` package

| Old Import | New Import |
|------------|------------|
| `from app_config import ...` | `from virelo.app.config import ...` |
| `from settings import Settings` | `from virelo.settings.persistence import Settings` |
| `from settings_state import SettingsState` | `from virelo.settings.state import SettingsState` |
| `from snap_service import SnapService` | `from virelo.services.snap import SnapService` |
| `from bridge import VireloBridge` | `from virelo.bridge.bridge import VireloBridge` |
| `from capture_guard import CaptureGuard` | `from virelo.bridge.capture_guard import CaptureGuard` |
| `from theme import ...` | `from virelo.services.theme import ...` |
| `from startup_shortcut import ...` | `from virelo.services.startup import ...` |
| `from workers import ...` | `from virelo.workers.key_capture import KeyCaptureWorker` / `from virelo.workers.explorer import ExplorerAutosizeWorker` |
| `from webview import VireloWebView` | `from virelo.app.webview import VireloWebView` (or keep at `virelo/` root) |
| `from explorer_columns import ...` | `from virelo.services.explorer_columns import ...` |

### Conditional PySide6 Import (for CI compatibility)

**Source:** `workers.py` lines 630-636
**Apply to:** `virelo/workers/key_capture.py`, `virelo/workers/explorer.py`, any module with QObject classes that unit tests may import on Linux CI

```python
try:
    from PySide6 import QtCore
except Exception:
    QtCore = None

if QtCore is not None:
    class KeyCaptureWorker(QtCore.QObject):
        # ...
```

### pyproject.toml Ruff per-file-ignores

**Apply to:** All `__init__.py` files with re-exports

```toml
[tool.ruff.lint.per-file-ignores]
"virelo/*/__init__.py" = ["F401"]
```

---

## No Analog Found

Files with no close match in the codebase (planner should use RESEARCH.md patterns instead):

| File | Role | Data Flow | Reason |
|------|------|-----------|--------|
| `pyproject.toml` | config | -- | No Python packaging config exists in the project. Use RESEARCH.md Code Examples (PEP 621 pattern). |
| `.github/workflows/ci.yml` | config | -- | No CI workflow exists. Use RESEARCH.md Code Examples (GitHub Actions pattern). |
| `frontend/src/test-setup.js` | config | -- | No test setup exists. Single line: `import '@testing-library/jest-dom';` |
| `tests/integration/conftest.py` | test | -- | No integration test infra exists. Use RESEARCH.md Pitfall 6 for `pytest.importorskip("PySide6")` pattern. |

---

## Metadata

**Analog search scope:** `D:\projects\Virelo\` (all Python and JavaScript source files)
**Files scanned:** 22 source files (12 Python root modules, 8 frontend JS/JSX files, 1 spec file, 1 vite config)
**Pattern extraction date:** 2026-04-24
