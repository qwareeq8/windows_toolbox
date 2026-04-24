# Phase 4: Snap and Explorer Hardening - Pattern Map

**Mapped:** 2026-04-24
**Files analyzed:** 4 (3 modified, 1 new)
**Analogs found:** 4 / 4

## File Classification

| New/Modified File | Role | Data Flow | Closest Analog | Match Quality |
|---|---|---|---|---|
| `virelo/services/snap.py` | service | event-driven | `virelo/services/snap.py` (self) | self — internal refactor |
| `virelo/services/explorer_service.py` | service | event-driven | `virelo/services/snap.py` (SnapService) | exact — same facade + lifecycle pattern |
| `virelo/app/window.py` | app-shell | request-response | `virelo/app/window.py` (self) | self — delegation extraction |
| `tests/unit/test_snap_geometry.py` | test | transform | `tests/unit/test_snap_geometry.py` (self) | self — extend existing file |

---

## Pattern Assignments

### `virelo/services/snap.py` — HotkeyListener extraction (D-01/D-02/D-03) + Virelo window exclusion (D-04/D-05/D-06)

**Analog:** `virelo/services/snap.py` (self — internal restructure, no external analog needed)

#### Imports pattern (lines 1-31)
```python
import ctypes
import logging
import threading
import time
from collections import deque
from ctypes import wintypes

import keyboard
import win32con
import win32gui
from PySide6 import QtCore

from virelo.app.config import normalize_snap_presses
from virelo.platform.win32_helpers import (
    USER32,
    _exit_fullscreen,
    _is_window_fullscreen,
    _should_skip_snap_for_game,
    get_monitor_rect,
)

LOG = logging.getLogger("Virelo")
```
No new imports required. HotkeyListener is a QObject subclass; the existing QtCore import covers `QtCore.QObject`, `QtCore.Signal`, and `QtCore.Slot`.

#### HotkeyListener: what to extract from ShiftSnapRestore.__init__ (lines 100-110)
These attributes and hooks move verbatim into HotkeyListener.__init__:
```python
self._press_times: deque[float] = deque(
    maxlen=normalize_snap_presses(self.settings.snap_presses)
)
self._press_lock = threading.Lock()
self._held = False
self.current_key = str(settings.snap_key)
self.restore_key = str(getattr(settings, "restore_key", "ctrl"))
self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)
```

#### HotkeyListener: what to extract — cleanup (lines 112-120)
```python
def cleanup(self):
    try:
        keyboard.unhook(self._press_hook)
    except Exception:
        pass
    try:
        keyboard.unhook(self._release_hook)
    except Exception:
        pass
```

#### HotkeyListener: what to extract — update_binding (lines 173-189)
```python
def update_binding(self, new_key: str):
    try:
        keyboard.unhook(self._press_hook)
    except Exception:
        pass
    try:
        keyboard.unhook(self._release_hook)
    except Exception:
        pass
    self.current_key = new_key
    self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
    self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)
    with self._press_lock:
        self._press_times = deque(
            self._press_times,
            maxlen=normalize_snap_presses(self.settings.snap_presses),
        )
```

#### HotkeyListener: what to extract — update_restore_key and update_press_limit (lines 191-196)
```python
def update_restore_key(self, new_key: str):
    self.restore_key = new_key

def update_press_limit(self, new_limit: int):
    with self._press_lock:
        self._press_times = deque(self._press_times, maxlen=new_limit)
```

#### HotkeyListener: what to extract — _on_press and _on_release (lines 198-222)
The signal emission `self.triggered.emit(restore)` stays here; `triggered` moves to HotkeyListener:
```python
def _on_press(self, event):
    if not self.settings.enable_snap:
        return
    if not self._held:
        self._held = True
        now = time.time()
        interval = int(self.settings.snap_interval) / 1000.0
        press_target = normalize_snap_presses(self.settings.snap_presses)
        should_trigger = False
        restore = False
        with self._press_lock:
            self._press_times = deque(
                [t for t in self._press_times if now - t <= interval],
                maxlen=press_target,
            )
            self._press_times.append(now)
            if len(self._press_times) >= press_target:
                self._press_times.clear()
                restore = keyboard.is_pressed(self.restore_key)
                should_trigger = True
        if should_trigger:
            self.triggered.emit(restore)

def _on_release(self, event):
    self._held = False
```

#### ShiftSnapRestore after split — what stays

`triggered = QtCore.Signal(bool)` moves to HotkeyListener. `blocked = QtCore.Signal(str)` stays on ShiftSnapRestore. `__init__` retains only:
```python
def __init__(self, settings):
    super().__init__()
    self.settings = settings
    self._orig_sizes: dict[int, dict[str, tuple[int, int, int, int] | bool]] = {}
    self._fetch_open_windows()
```
`_fetch_open_windows`, `_prune_closed_windows`, `perform`, `_snap`, `_restore` are unchanged in logic.

#### Virelo window exclusion in _snap (lines 238-248) — D-04

Current behavior to replace:
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        center_fn = getattr(widget, "center_on_screen", None)
        if center_fn is not None:
            center_fn()
            widget.raise_()
            widget.activateWindow()
        return
```

New behavior (D-04/D-06):
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        LOG.debug("Snap: skipping Virelo's own window hwnd=%s", hwnd)
        return  # Skip entirely per SNAP-03
```

#### Virelo window exclusion in _restore (lines 316-329) — D-05

Current behavior to replace (Qt branch of _restore):
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        orig = self._orig_sizes.pop(hwnd, None)
        if not orig:
            return
        rect = orig["rect"] if isinstance(orig, dict) else orig
        left, top, width, height = rect
        widget.setGeometry(left, top, width, height)
        widget.raise_()
        widget.activateWindow()
        return
```

New behavior (D-05/D-06):
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        LOG.debug("Restore: skipping Virelo's own window hwnd=%s", hwnd)
        return  # Skip entirely per D-05
```

#### SnapService delegation after split — D-02/Pitfall 2

SnapService currently delegates `update_binding`, `update_restore_key`, and `update_press_limit` to `self._mgr` (ShiftSnapRestore, lines 72-85). After the split, these three methods live on HotkeyListener. The cleanest option (RESEARCH A3) is to hold a second reference on SnapService:

```python
class SnapService:
    def __init__(self, shift_mgr, hotkey_listener=None):
        self._mgr = shift_mgr
        self._listener = hotkey_listener  # NEW

    def set_manager(self, mgr):
        self._mgr = mgr

    def set_listener(self, listener):   # NEW
        self._listener = listener

    def update_binding(self, key: str):
        if self._listener:
            self._listener.update_binding(key)

    def update_restore_key(self, key: str):
        if self._listener:
            self._listener.update_restore_key(key)

    def update_press_limit(self, count: int):
        if self._listener:
            self._listener.update_press_limit(count)

    # test_snap still delegates to self._mgr.perform unchanged
```

---

### `virelo/services/explorer_service.py` — NEW file (D-07/D-08/D-09/D-10)

**Analog:** `virelo/services/snap.py` — SnapService facade (lines 50-86)

This is the closest match: same pattern of a plain class (not QObject) that wraps a QThread + QObject worker, exposes start/stop lifecycle, and gates on a settings flag.

#### Imports pattern — derived from window.py module-level (lines 1-29) and snap.py (lines 1-31)
```python
import logging
import time

from virelo.platform.win32_helpers import _is_window_interactive
from virelo.services.explorer_columns import autosize_explorer_columns
from virelo.workers.explorer import ExplorerAutosizeWorker

LOG = logging.getLogger("Virelo")
```
`QtCore.QThread` is imported lazily inside `start()` to mirror the worker's conditional PySide6 import pattern (`workers/explorer.py` lines 624-628).

#### Core pattern — SnapService facade structure (snap.py lines 50-86)
SnapService is the structural template:
```python
class SnapService:
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

ExplorerService follows this skeleton — plain class, settings-gated, start/stop/is_running, no public internals.

#### Worker start pattern — window.py _update_explorer_autosize_thread (lines 424-469)
This entire method body becomes `ExplorerService.start()`. Extract verbatim, substituting `self._settings` for `self.settings` and `self._thread`/`self._worker` for the instance attributes:
```python
def _update_explorer_autosize_thread(self, *args):
    LOG.info("_update_explorer_autosize_thread: called")
    app = QtWidgets.QApplication.instance()
    pushed_cursor = False
    if app is not None:
        QtGui.QGuiApplication.setOverrideCursor(
            QtGui.QCursor(QtCore.Qt.CursorShape.WaitCursor)
        )
        pushed_cursor = True
    try:
        group_enabled = bool(self.settings.ex_auto_size)
        if not group_enabled:
            self._stop_explorer_worker()
            return
        if self._explorer_thread and self._explorer_thread.isRunning():
            return
        self._explorer_thread = QtCore.QThread(self)
        self._explorer_worker = ExplorerAutosizeWorker(
            _autosize_explorer_columns_quick,
            _autosize_explorer_columns_full,
            _is_window_interactive,
            schedule=(0.05, 0.1, 0.25, 0.5, 1.0),
        )
        self._explorer_worker.moveToThread(self._explorer_thread)
        self._explorer_thread.started.connect(self._explorer_worker.run)
        self._explorer_worker.finished.connect(self._explorer_thread.quit)
        self._explorer_worker.finished.connect(self._explorer_worker.deleteLater)
        self._explorer_thread.finished.connect(self._explorer_thread.deleteLater)
        self._explorer_thread.finished.connect(self._on_explorer_finished)
        self._explorer_thread.start()
    finally:
        if pushed_cursor:
            QtGui.QGuiApplication.restoreOverrideCursor()
```

#### Worker stop pattern — window.py _stop_explorer_worker (lines 474-492)
This becomes `ExplorerService.stop()`. Extract verbatim:
```python
def _stop_explorer_worker(self):
    worker = getattr(self, "_explorer_worker", None)
    thread = getattr(self, "_explorer_thread", None)
    if worker is not None:
        try:
            worker.stop()
        except Exception:
            pass
        time.sleep(0.05)  # Give worker time to see stop flag before thread.quit
    if thread is not None:
        thread.quit()
        if not thread.wait(3000):
            LOG.warning("Explorer autosize: thread did not stop in time")
    self._explorer_worker = None
    self._explorer_thread = None
    LOG.info("Explorer autosize: worker stopped.")
```

#### Finished callback — window.py _on_explorer_finished (lines 494-496)
```python
def _on_explorer_finished(self):
    self._explorer_worker = None
    self._explorer_thread = None
```

#### Module-level wrapper functions — window.py lines 40-67
These two functions move from window.py into explorer_service.py unchanged. They are passed as callables to ExplorerAutosizeWorker:
```python
def _autosize_explorer_columns_quick(
    top_hwnd: int, target_path: str = None, caller_owns_com: bool = False
) -> tuple:
    return autosize_explorer_columns(
        top_hwnd,
        allow_keyboard_fallback=False,
        target_path=target_path,
        caller_owns_com=caller_owns_com,
    )

def _autosize_explorer_columns_full(
    top_hwnd: int, target_path: str = None, caller_owns_com: bool = False
) -> tuple:
    return autosize_explorer_columns(
        top_hwnd,
        allow_keyboard_fallback=False,
        target_path=target_path,
        caller_owns_com=caller_owns_com,
    )
```

#### COM constraint — workers/explorer.py docstring (lines 1-11)
ExplorerService MUST NOT touch any `pythoncom`, `win32com.client`, or `Shell.Application` directly. The constraint from the worker's module docstring applies:
> COM init, Shell.Application caching, iter_tabs, and all COM-dependent closures MUST stay in this single file. Do NOT separate COM init from COM usage.

---

### `virelo/app/window.py` — ExplorerService wiring + SnapService listener wiring (D-08/Pitfall 3/Pitfall 4)

**Analog:** `virelo/app/window.py` (self — targeted delegation updates)

#### Import addition
Add to the imports block (after line 25 `from virelo.services.snap import ShiftSnapRestore, SnapService`):
```python
from virelo.services.explorer_service import ExplorerService
```
Remove the now-redundant `from virelo.workers.explorer import ExplorerAutosizeWorker` import (line 27) — ExplorerService owns that import.
Remove the now-redundant `_autosize_explorer_columns_quick` and `_autosize_explorer_columns_full` module-level functions (lines 40-67) — these move to explorer_service.py.

#### __init__ changes — ExplorerService construction and HotkeyListener wiring (lines 241-249)

Current wiring (lines 241-249):
```python
self.shift_mgr = ShiftSnapRestore(self.settings)
self.shift_mgr.triggered.connect(self.shift_mgr.perform)
self.shift_mgr.blocked.connect(lambda message: self.snap_key_status.emit(message, 3000))
self._snap_service.set_manager(self.shift_mgr)
self._update_explorer_enabled_state()
```

New wiring after split — add HotkeyListener and ExplorerService:
```python
# HotkeyListener + ShiftSnapRestore construction
self._hotkey_listener = HotkeyListener(self.settings)
self.shift_mgr = ShiftSnapRestore(self.settings)
self._hotkey_listener.triggered.connect(self.shift_mgr.perform)  # cross-object signal wiring
self.shift_mgr.blocked.connect(lambda message: self.snap_key_status.emit(message, 3000))
self._snap_service.set_manager(self.shift_mgr)
self._snap_service.set_listener(self._hotkey_listener)  # NEW

# ExplorerService replaces inline thread management
self._explorer_service = ExplorerService(self.settings, parent=self)
self._update_explorer_enabled_state()  # still calls _explorer_service.start()
```

#### _update_explorer_enabled_state — thin delegate (lines 421-422)

Current:
```python
def _update_explorer_enabled_state(self):
    self._update_explorer_autosize_thread()
```

After extraction:
```python
def _update_explorer_enabled_state(self):
    self._explorer_service.start()
```

#### _update_explorer_autosize_thread — kept as thin delegate to avoid breaking bridge.py (lines 424-472 / Pitfall 3)

`bridge.py` line 271 calls `mw._update_explorer_autosize_thread()` and `bridge.py` line 135 calls the same. Keep a thin shim:
```python
def _update_explorer_autosize_thread(self, *args):
    self._explorer_service.start()
```

#### _stop_background_threads — update cleanup path (lines 289-292 / Pitfall 4)

Current:
```python
def _stop_background_threads(self):
    self._stop_capture_worker()
    self._stop_explorer_worker()
    self._stop_theme_sync()
```

After extraction:
```python
def _stop_background_threads(self):
    self._stop_capture_worker()
    self._explorer_service.stop()
    self._stop_theme_sync()
```

#### shift_mgr.cleanup() call in closeEvent and _really_quit — extend to hotkey_listener
`closeEvent` (line 269) and `_really_quit` (line 287) both call `self.shift_mgr.cleanup()`. After the split, `cleanup()` moves to HotkeyListener. Update both callsites:
```python
self._hotkey_listener.cleanup()  # replaces self.shift_mgr.cleanup()
```

#### Key capture callbacks — update_binding / update_restore_key delegation (lines 329-343)
These currently call `self.shift_mgr.update_binding(key_str)` and `self.shift_mgr.update_restore_key(key_str)`. After the split, these delegate through SnapService which now holds the listener reference — so `self._snap_service.update_binding(key_str)` and `self._snap_service.update_restore_key(key_str)` are the clean call path, or directly `self._hotkey_listener.update_binding(key_str)`. Prefer the SnapService path for consistency.

#### _reset_defaults — shift_mgr.update_* calls (lines 394-396)
Same as above — these delegate through SnapService or _hotkey_listener directly:
```python
# Current (lines 394-396):
if hasattr(self, "shift_mgr"):
    self.shift_mgr.update_binding(defaults["snap_key"])
    self.shift_mgr.update_restore_key(defaults["restore_key"])
    self.shift_mgr.update_press_limit(defaults["snap_presses"])

# After split — delegate to hotkey_listener (or via snap_service):
if hasattr(self, "_hotkey_listener"):
    self._hotkey_listener.update_binding(defaults["snap_key"])
    self._hotkey_listener.update_restore_key(defaults["restore_key"])
    self._hotkey_listener.update_press_limit(defaults["snap_presses"])
```

---

### `tests/unit/test_snap_geometry.py` — extend with D-12/D-13/D-14 (D-15)

**Analog:** `tests/unit/test_snap_geometry.py` (self — extend existing file)

#### Existing file structure (lines 1-68)
All three new test functions must follow the same pattern as the existing `calculate_snap_position` tests:
- Module-level functions (no class wrapping)
- Descriptive docstring as first line
- Import only from `virelo.services.snap` and `virelo.platform.win32_helpers`
- No fixtures needed for pure geometry tests

#### Existing test pattern to copy — calculate_snap_position style (lines 37-44)
```python
def test_calculate_snap_position_76pct():
    """76% width/height on 1920x1080 starting at (0,0)."""
    x, y, w, h = calculate_snap_position(0, 0, 1920, 1080, 76, 76)
    assert w == 1920 * 76 // 100  # 1459
    assert h == 1080 * 76 // 100  # 820
    assert x == (1920 - w) // 2  # 230
    assert y == (1080 - h) // 2  # 130
```

#### D-12 — negative coords test (RESEARCH Code Examples)
```python
def test_calculate_snap_position_negative_coords():
    """Snap geometry on monitor with negative origin (left-of-primary)."""
    x, y, w, h = calculate_snap_position(-1920, 0, 1920, 1080, 76, 76)
    assert w == 1920 * 76 // 100  # 1459
    assert h == 1080 * 76 // 100  # 820
    assert x == -1920 + (1920 - w) // 2  # -1690
    assert y == (1080 - h) // 2  # 130
```

#### D-13 — vertical layout test (RESEARCH Code Examples)
```python
def test_calculate_snap_position_vertical_layout():
    """Snap on monitor below primary (y offset, vertical multi-monitor)."""
    x, y, w, h = calculate_snap_position(0, 1080, 2560, 1440, 76, 76)
    assert w == 2560 * 76 // 100  # 1945
    assert h == 1440 * 76 // 100  # 1094
    assert x == (2560 - w) // 2  # 307
    assert y == 1080 + (1440 - h) // 2  # 1253
```

#### D-14 — maximized restore test

The restore path needs ShiftSnapRestore. After the split, ShiftSnapRestore no longer subclasses QObject in the normal sense — but `conftest.py` stubs `PySide6.QtCore.QObject` as `_StubQObject` (lines 30-39) and `Signal`/`Slot` as MagicMocks (lines 40-41), so `ShiftSnapRestore.__new__` instantiation works in unit tests.

The `win32gui.ShowWindow` stub is already a `MagicMock()` (conftest.py line 73), so `assert_called_with` works immediately.

Test pattern to follow — uses `__new__` to bypass `__init__` (matches RESEARCH D-14 example):
```python
def test_restore_maximized_window():
    """Restore of a previously-maximized window issues SW_MAXIMIZE."""
    import win32con
    import win32gui

    from virelo.services.snap import ShiftSnapRestore

    mgr = ShiftSnapRestore.__new__(ShiftSnapRestore)
    mgr.settings = MockSettings()
    mgr._orig_sizes = {
        12345: {
            "rect": (100, 100, 800, 600),
            "maximized": True,
        }
    }
    # Stub GetForegroundWindow / GetWindowRect / get_monitor_rect
    # to control flow into the _restore non-Qt branch
    # win32gui.ShowWindow is already a MagicMock from conftest
```

To reach the `was_maximized` branch in `_restore` (snap.py lines 355-356), the test must ensure:
1. `hwnd` is NOT in `QtWidgets.QApplication.topLevelWidgets()` — the `PySide6.QtWidgets` stub is a bare module with no `QApplication` attribute, so the for-loop is empty by default. Confirm or stub accordingly.
2. `USER32.GetWindowRect` returns a rect that does NOT cover the monitor (so the early return on lines 347-353 is skipped).
3. `get_monitor_rect(hwnd)` returns a valid tuple.

These require patching `virelo.services.snap.get_monitor_rect` and `virelo.services.snap.USER32.GetWindowRect` with `unittest.mock.patch`.

The `MockSettings` fixture is available from conftest (line 89-105) — import it or use the `mock_settings` fixture.

---

## Shared Patterns

### Logging — apply to all modified files
**Source:** `virelo/services/snap.py` line 31 and `virelo/workers/explorer.py` lines 697-699
```python
LOG = logging.getLogger("Virelo")
```
All new and modified Python files use this exact logger name. Debug messages use `LOG.debug(...)`, info uses `LOG.info(...)`, warnings use `LOG.warning(...)`. The pattern for skip-with-reason messages (D-06) follows the existing game-mode skip at snap.py line 274:
```python
LOG.info("Game mode: skipped snap for fullscreen window hwnd=%s", hwnd)
```

### QThread lifecycle — apply to ExplorerService.start() and ExplorerService.stop()
**Source:** `virelo/app/window.py` lines 449-465 (start), lines 474-492 (stop)

The canonical QThread wiring sequence in this codebase is:
```python
thread = QtCore.QThread(parent)
worker.moveToThread(thread)
thread.started.connect(worker.run)
worker.finished.connect(thread.quit)
worker.finished.connect(worker.deleteLater)
thread.finished.connect(thread.deleteLater)
thread.finished.connect(self._on_finished)  # cleanup callback
thread.start()
```
The stop sequence always calls `worker.stop()` first, waits `0.05s` for the stop flag to propagate, then `thread.quit()` + `thread.wait(3000)`.

### Conftest stub usage — apply to D-14 restore tests
**Source:** `tests/conftest.py` lines 46-77

All win32 stubs are `MagicMock()` instances. Assertions use standard mock API:
```python
win32gui.ShowWindow.assert_called_with(hwnd, win32con.SW_MAXIMIZE)
```
No new stubs needed. `MockSettings` (conftest lines 89-105) provides the settings stand-in for `ShiftSnapRestore.__new__` tests.

### Error handling — apply to ExplorerService
**Source:** `virelo/services/snap.py` lines 62-70 (SnapService.test_snap)
```python
try:
    self._mgr.perform(False)
    return {"ok": True, "message": "..."}
except Exception as e:
    LOG.exception("test_snap failed")
    return {"ok": False, "error": str(e)}
```
Bare `except Exception` + `LOG.exception` is the project convention for service-layer error boundaries. Worker lifecycle errors in `stop()` use `try/except Exception: pass` (see window.py lines 478-480) — the same pattern used for `cleanup()` in snap.py lines 113-120.

---

## No Analog Found

All files in scope have close analogs in the codebase. No files require falling back to RESEARCH.md patterns exclusively.

---

## Metadata

**Analog search scope:** `virelo/services/`, `virelo/app/`, `virelo/workers/`, `virelo/bridge/`, `tests/unit/`, `tests/conftest.py`
**Files scanned:** 8 primary source files read in full
**Pattern extraction date:** 2026-04-24
