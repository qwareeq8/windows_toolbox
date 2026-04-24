#!/usr/bin/env python
# virelo.py
# PySide6-based Virelo. Manage snap and restore with minimize-to-tray and close-to-tray behavior.

import ctypes
import logging
import os
import sys
import threading
import time
from collections import deque
from ctypes import wintypes
from logging.handlers import RotatingFileHandler

# Exit early on non-Windows platforms.
if sys.platform != "win32":
    print("Virelo requires Windows.")
    sys.exit(1)

import atexit
import faulthandler

import keyboard
import win32api
import win32con
import win32event
import win32gui
import winerror
from PySide6 import QtCore, QtGui, QtWidgets
from win32com.client import Dispatch

from app_config import (
    APP_ID,
    APP_NAME,
    DEFAULTS,
    LOG_DIR,
    LOG_FILE,
    ORGANIZATION,
    normalize_snap_presses,
)
from bridge import VireloBridge
from capture_guard import CaptureGuard
from explorer_columns import (
    autosize_explorer_columns,
)
from settings import Settings
from settings_state import SettingsState
from snap_service import SnapService
from startup_shortcut import startup_shortcut_spec
from theme import get_windows_theme, normalize_theme_mode, resolve_theme, toggle_theme_mode
from webview import VireloWebView
from workers import ExplorerAutosizeWorker, KeyCaptureWorker

# ------------------------------------------------------------------------------
# Logging and crash diagnostics
# ------------------------------------------------------------------------------


def _init_logger() -> logging.Logger:
    """Initialize a rotating file logger in a per-user location."""
    base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    log_dir = os.path.join(base, LOG_DIR)
    os.makedirs(log_dir, exist_ok=True)

    log_path = os.path.join(log_dir, LOG_FILE)
    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.DEBUG)  # Enable DEBUG level by default for troubleshooting
    logger.propagate = False

    existing_handler = None
    for handler in logger.handlers:
        if isinstance(handler, RotatingFileHandler) and getattr(
            handler, "baseFilename", None
        ) == os.path.abspath(log_path):
            existing_handler = handler
            break

    if existing_handler is None:
        handler = RotatingFileHandler(
            log_path,
            maxBytes=512 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        handler.setLevel(logging.DEBUG)  # File handler captures all levels
        logger.addHandler(handler)
        existing_handler = handler

    # Also add a console handler for immediate feedback during development
    console_handler = None
    for h in logger.handlers:
        if isinstance(h, logging.StreamHandler) and not isinstance(h, RotatingFileHandler):
            console_handler = h
            break

    if console_handler is None:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        console_handler.setLevel(logging.INFO)  # Console shows INFO and above
        logger.addHandler(console_handler)

    logger.log_path = getattr(existing_handler, "baseFilename", log_path)
    return logger


LOG = _init_logger()

_CRASH_LOG = None
try:
    crash_log_path = os.path.join(os.path.dirname(getattr(LOG, "log_path", "")), "crash.log")
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


def _ensure_dispatch(app_name: str):
    try:
        return Dispatch(app_name)
    except AttributeError:
        import re
        import shutil

        LOG.warning("win32com gen_py cache appears corrupted. Rebuilding.")
        module_list = [m.__name__ for m in sys.modules.values() if getattr(m, "__name__", None)]
        for module in module_list:
            if re.match(r"win32com\.gen_py\..+", module):
                sys.modules.pop(module, None)
        localappdata = os.environ.get("LOCALAPPDATA")
        if localappdata:
            gen_py_path = os.path.join(localappdata, "Temp", "gen_py")
            if os.path.exists(gen_py_path):
                shutil.rmtree(gen_py_path, ignore_errors=True)
        from win32com import client

        return client.gencache.EnsureDispatch(app_name)


USER32 = ctypes.windll.user32
KERNEL32 = ctypes.windll.kernel32

LVM_FIRST = 0x1000
LVM_GETHEADER = LVM_FIRST + 31
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_SETCOLUMNWIDTH = LVM_FIRST + 30

LVSCW_AUTOSIZE = -1
LVSCW_AUTOSIZE_USEHEADER = -2

HDM_FIRST = 0x1200
HDM_GETITEMCOUNT = HDM_FIRST + 0

WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_LBUTTONDBLCLK = 0x0203
SMTO_ABORTIFHUNG = 0x0002

UIA_BoundingRectanglePropertyId = 30001
UIA_ControlTypePropertyId = 30003
UIA_NativeWindowHandlePropertyId = 30020
UIA_HeaderItemControlTypeId = 50035

TreeScope_Children = 2
TreeScope_Subtree = 7

APP_TITLE = APP_NAME

INITIAL_SIZE = QtCore.QSize(640, 520)
MIN_SIZE = QtCore.QSize(560, 420)


def resource_path(relative_path: str) -> str:
    """Return absolute path to resource, works for dev and PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def get_monitor_rect(hwnd: int, use_work_area: bool = True) -> tuple[int, int, int, int] | None:
    """
    Get monitor rectangle for a given window.

    Args:
        hwnd: Window handle
        use_work_area: If True, return work area (taskbar-adjusted).
                      If False, return full monitor bounds (for fullscreen detection).

    Returns:
        (left, top, right, bottom) tuple or None
    """
    try:
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        info = win32api.GetMonitorInfo(monitor)
        if use_work_area:
            # Prefer work area for normal snapping (respects taskbar)
            return info.get("Work") or info.get("Monitor")
        else:
            # Prefer full monitor for fullscreen detection (ignores taskbar)
            return info.get("Monitor") or info.get("Work")
    except Exception:
        return None


def _get_window_dwm_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """
    Get window rect using DWM extended frame bounds if available.

    DWM extended frame bounds exclude invisible borders and give the true
    visual bounds of the window, which is more accurate for fullscreen detection.

    Falls back to GetWindowRect if DWM attributes are not available.

    Returns:
        (left, top, right, bottom) tuple or None
    """
    try:
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        rect = wintypes.RECT()
        result = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect), ctypes.sizeof(rect)
        )
        if result == 0:
            return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass
    # Fallback to standard GetWindowRect
    return _get_rect(hwnd)


FULLSCREEN_TOLERANCE = 3


def _rect_matches_monitor(
    rect: tuple[int, int, int, int], monitor: tuple[int, int, int, int]
) -> bool:
    left, top, right, bottom = rect
    left_edge, top_edge, right_edge, bottom_edge = monitor
    return (
        abs(left - left_edge) <= FULLSCREEN_TOLERANCE
        and abs(top - top_edge) <= FULLSCREEN_TOLERANCE
        and abs(right - right_edge) <= FULLSCREEN_TOLERANCE
        and abs(bottom - bottom_edge) <= FULLSCREEN_TOLERANCE
    )


def _is_window_fullscreen(
    hwnd: int,
    rect: wintypes.RECT | None = None,
    monitor_rect: tuple[int, int, int, int] | None = None,
) -> bool:
    """
    Check if window is fullscreen using true monitor bounds.

    Uses DWM extended frame bounds for accurate window rect,
    and full monitor bounds (not work area) for comparison.

    Args:
        hwnd: Window handle
        rect: Optional pre-fetched window rect (for optimization)
        monitor_rect: Optional pre-fetched monitor rect (must be FULL monitor bounds)

    Returns:
        True if window appears to be fullscreen
    """
    try:
        if not win32gui.IsWindow(hwnd):
            return False
        if monitor_rect is None:
            # IMPORTANT: Use full monitor bounds, not work area
            monitor_rect = get_monitor_rect(hwnd, use_work_area=False)
        if monitor_rect is None:
            return False

        # Try DWM extended frame bounds first for accuracy
        if rect is None:
            dwm_rect = _get_window_dwm_rect(hwnd)
            if dwm_rect:
                return _rect_matches_monitor(dwm_rect, monitor_rect)
            # Fallback to standard GetWindowRect
            rect = wintypes.RECT()
            USER32.GetWindowRect(hwnd, ctypes.byref(rect))

        rect_vals = (rect.left, rect.top, rect.right, rect.bottom)
        return _rect_matches_monitor(rect_vals, monitor_rect)
    except Exception:
        return False


def _looks_like_game_window(hwnd: int) -> bool:
    try:
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    except Exception:
        return False
    has_caption = bool(style & win32con.WS_CAPTION or style & win32con.WS_BORDER)
    is_popup = bool(style & win32con.WS_POPUP)
    return is_popup and not has_caption


def _should_skip_snap_for_game(hwnd: int, settings: Settings, full_screen: bool) -> bool:
    return (
        full_screen
        and getattr(settings, "game_mode_enabled", True)
        and _looks_like_game_window(hwnd)
    )


def _exit_fullscreen(hwnd: int):
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] in (win32con.SW_MAXIMIZE, win32con.SW_SHOWMAXIMIZED):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            return
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if not style & win32con.WS_CAPTION:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass


def _as_hwnd(v: int) -> int:
    try:
        return int(ctypes.c_size_t(int(v)).value or 0)
    except Exception:
        return 0


def _get_children(hwnd: int):
    try:
        child = win32gui.GetWindow(hwnd, win32con.GW_CHILD)
        while child:
            yield child
            child = win32gui.GetWindow(child, win32con.GW_HWNDNEXT)
    except Exception:
        LOG.debug("GetWindow failed for hwnd=%s", hwnd, exc_info=True)
        return


def _class_name(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd)
    except Exception:
        return ""


def _get_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    try:
        rc = wintypes.RECT()
        if USER32.GetWindowRect(hwnd, ctypes.byref(rc)):
            return rc.left, rc.top, rc.right, rc.bottom
    except Exception:
        pass
    return None


def _area(hwnd: int) -> int:
    r = _get_rect(hwnd)
    if not r:
        return 0
    left_edge, top_edge, right_edge, bottom_edge = r
    return max(0, right_edge - left_edge) * max(0, bottom_edge - top_edge)


def _ancestor_classes(hwnd: int, depth: int = 8) -> tuple[str, ...]:
    out = []
    try:
        cur = hwnd
        for _ in range(depth):
            cur = win32gui.GetParent(cur)
            if not cur:
                break
            out.append(_class_name(cur))
    except Exception:
        pass
    return tuple(out)


def _find_descendant_by_class(
    hwnd_start: int, class_names: tuple, max_depth: int = 12
) -> int | None:
    try:
        class_names = tuple(n.lower() for n in class_names)
        queue = [(hwnd_start, 0)]
        visited = set()
        while queue:
            hwnd, depth = queue.pop(0)
            if hwnd in visited or depth > max_depth:
                continue
            visited.add(hwnd)
            try:
                cname = win32gui.GetClassName(hwnd).lower()
            except Exception:
                cname = ""
            if cname in class_names and hwnd != hwnd_start:
                return hwnd
            for ch in _get_children(hwnd):
                queue.append((ch, depth + 1))
    except Exception as e:
        LOG.exception("find_descendant_by_class failed", exc_info=e)
    return None


def _collect_descendants_by_class(
    hwnd_start: int, class_names: tuple, max_depth: int = 12
) -> tuple[int, ...]:
    found = []
    try:
        class_names = tuple(n.lower() for n in class_names)
        queue = [(hwnd_start, 0)]
        visited = set()
        while queue:
            hwnd, depth = queue.pop(0)
            if hwnd in visited or depth > max_depth:
                continue
            visited.add(hwnd)
            try:
                cname = win32gui.GetClassName(hwnd).lower()
            except Exception:
                cname = ""
            if cname in class_names and hwnd != hwnd_start:
                found.append(hwnd)
            for ch in _get_children(hwnd):
                queue.append((ch, depth + 1))
    except Exception as e:
        LOG.exception("collect_descendants_by_class failed", exc_info=e)
    return tuple(found)


def _is_window_interactive(hwnd: int) -> bool:
    try:
        if not win32gui.IsWindow(hwnd):
            return False
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if win32gui.IsIconic(hwnd):
            return False
        rc = wintypes.RECT()
        USER32.GetWindowRect(hwnd, ctypes.byref(rc))
        return (rc.right - rc.left) > 0 and (rc.bottom - rc.top) > 0
    except Exception:
        return False


def _looks_like_preview(hwnd: int) -> bool:
    """Heuristic: any ancestor class mentions 'preview'."""
    try:
        for cls in _ancestor_classes(hwnd, depth=10):
            if "preview" in (cls or "").lower():
                return True
    except Exception:
        pass
    return False


def _find_best_folder_listview(top_hwnd: int) -> int | None:
    """
    Prefer the FolderView listview under SHELLDLL_DefView.
    If multiple SysListView32 exist (e.g., Preview pane), choose the largest non-preview one.
    """
    defview = _find_descendant_by_class(top_hwnd, ("SHELLDLL_DefView",), max_depth=12)
    candidates = []
    if defview:
        candidates = list(_collect_descendants_by_class(defview, ("SysListView32",), max_depth=6))
    if not candidates:
        candidates = list(_collect_descendants_by_class(top_hwnd, ("SysListView32",), max_depth=14))
    if not candidates:
        return None

    filtered = [h for h in candidates if _is_window_interactive(h) and not _looks_like_preview(h)]
    if not filtered:
        filtered = [h for h in candidates if _is_window_interactive(h)]
    if not filtered:
        filtered = candidates

    best = max(filtered, key=_area, default=None)
    return _as_hwnd(best) if best else None


def _autosize_explorer_columns_quick(
    top_hwnd: int, target_path: str = None, caller_owns_com: bool = False
) -> tuple:
    """
    Single autosize attempt using COM-based column manager only.
    Returns (success, method).

    Args:
        top_hwnd: Top-level Explorer window handle
        target_path: If provided, find the tab matching this path (for Windows 11 tabs)
        caller_owns_com: If True, caller manages COM init/uninit
    """
    return autosize_explorer_columns(
        top_hwnd,
        allow_keyboard_fallback=False,
        target_path=target_path,
        caller_owns_com=caller_owns_com,
    )


def _autosize_explorer_columns_full(
    top_hwnd: int, target_path: str = None, caller_owns_com: bool = False
) -> tuple:
    """
    Full autosize attempt; currently identical to quick (COM-only, no fallbacks).
    Returns (success, method).

    Args:
        top_hwnd: Top-level Explorer window handle
        target_path: If provided, find the tab matching this path (for Windows 11 tabs)
        caller_owns_com: If True, caller manages COM init/uninit
    """
    return autosize_explorer_columns(
        top_hwnd,
        allow_keyboard_fallback=False,
        target_path=target_path,
        caller_owns_com=caller_owns_com,
    )


# Keep the old function name for backward compatibility with tests
def _autosize_explorer_columns_try(top_hwnd: int) -> bool:
    """
    Legacy wrapper for backward compatibility.

    Returns True if autosize succeeded, False otherwise.
    """
    success, method = _autosize_explorer_columns_quick(top_hwnd)
    return success


def get_startup_shortcut_path() -> str:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise RuntimeError("APPDATA is not set.")
    startup_dir = os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Startup")
    return os.path.join(startup_dir, f"{APP_NAME}.lnk")


def create_startup_shortcut():
    shortcut_path = get_startup_shortcut_path()
    script = os.path.abspath(sys.argv[0])
    frozen = bool(getattr(sys, "frozen", False))
    target, args = startup_shortcut_spec(sys.executable, script, frozen)
    wsh = Dispatch("WScript.Shell")
    os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)
    shortcut = wsh.CreateShortcut(shortcut_path)
    shortcut.TargetPath = target
    shortcut.Arguments = args
    shortcut.WorkingDirectory = os.path.dirname(target if frozen else script)
    icon_path = resource_path("icon.ico")
    if os.path.exists(icon_path):
        shortcut.IconLocation = icon_path
    shortcut.Save()


def remove_startup_shortcut():
    shortcut_path = get_startup_shortcut_path()
    if os.path.exists(shortcut_path):
        try:
            os.remove(shortcut_path)
        except Exception as e:
            LOG.exception("Failed to remove startup shortcut.", exc_info=e)


# ------------------------------------------------------------------------------
# SHIFT triple-press snap and restore
# ------------------------------------------------------------------------------


class ShiftSnapRestore(QtCore.QObject):
    triggered = QtCore.Signal(bool)
    blocked = QtCore.Signal(str)

    def __init__(self, settings: Settings):
        super().__init__()
        self.settings = settings
        self._press_times: deque[float] = deque(
            maxlen=normalize_snap_presses(self.settings.snap_presses)
        )
        self._press_lock = threading.Lock()
        self._held = False
        self._orig_sizes: dict[int, dict[str, tuple[int, int, int, int] | bool]] = {}
        self.current_key = str(settings.snap_key)
        self.restore_key = str(getattr(settings, "restore_key", "ctrl"))
        self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
        self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)
        self._fetch_open_windows()

    def cleanup(self):
        try:
            keyboard.unhook(self._press_hook)
        except Exception:
            pass
        try:
            keyboard.unhook(self._release_hook)
        except Exception:
            pass

    def _fetch_open_windows(self):
        def enum_windows_callback(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            try:
                placement = win32gui.GetWindowPlacement(hwnd)
                rc = wintypes.RECT()
                USER32.GetWindowRect(hwnd, ctypes.byref(rc))
                if placement[1] == win32con.SW_MAXIMIZE:
                    self._orig_sizes[hwnd] = {
                        "rect": (
                            rc.left,
                            rc.top,
                            rc.right - rc.left,
                            rc.bottom - rc.top,
                        ),
                        "maximized": True,
                    }
                else:
                    if rc.right - rc.left > 0 and rc.bottom - rc.top > 0:
                        self._orig_sizes[hwnd] = {
                            "rect": (
                                rc.left,
                                rc.top,
                                rc.right - rc.left,
                                rc.bottom - rc.top,
                            ),
                            "maximized": False,
                        }
            except Exception as e:
                LOG.exception("EnumWindows callback failed.", exc_info=e)
            return True

        self._orig_sizes.clear()
        win32gui.EnumWindows(enum_windows_callback, None)

    def _prune_closed_windows(self):
        existing = set()

        def enum_cb(hwnd, _):
            existing.add(hwnd)
            return True

        win32gui.EnumWindows(enum_cb, None)
        stale = [hwnd for hwnd in list(self._orig_sizes.keys()) if hwnd not in existing]
        for hwnd in stale:
            self._orig_sizes.pop(hwnd, None)

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

    def update_restore_key(self, new_key: str):
        self.restore_key = new_key

    def update_press_limit(self, new_limit: int):
        with self._press_lock:
            self._press_times = deque(self._press_times, maxlen=new_limit)

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

    @QtCore.Slot(bool)
    def perform(self, restore: bool):
        self._prune_closed_windows()
        hwnd = USER32.GetForegroundWindow()
        if not hwnd:
            return
        try:
            if restore:
                self._restore(hwnd)
            else:
                self._snap(hwnd)
        except Exception as e:
            LOG.exception("ShiftSnapRestore.perform failed.", exc_info=e)

    def _snap(self, hwnd: int):
        from PySide6 import QtWidgets

        for widget in QtWidgets.QApplication.topLevelWidgets():
            if int(widget.winId()) == hwnd:
                center_fn = getattr(widget, "center_on_screen", None)
                if center_fn is not None:
                    center_fn()
                    widget.raise_()
                    widget.activateWindow()
                return

        def refresh_rect():
            rect = wintypes.RECT()
            USER32.GetWindowRect(hwnd, ctypes.byref(rect))
            return rect

        rc = refresh_rect()
        if hwnd not in self._orig_sizes:
            placement = win32gui.GetWindowPlacement(hwnd)
            was_maximized = placement[1] == win32con.SW_MAXIMIZE
            self._orig_sizes[hwnd] = {
                "rect": (rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top),
                "maximized": was_maximized,
            }

        # Get full monitor bounds for accurate fullscreen detection
        mon_full = get_monitor_rect(hwnd, use_work_area=False)
        if not mon_full:
            return

        # Check if window is fullscreen using full monitor bounds
        full_screen = _is_window_fullscreen(hwnd, rect=rc, monitor_rect=mon_full)

        # Skip snapping if game mode enabled and window is fullscreen borderless
        if _should_skip_snap_for_game(hwnd, self.settings, full_screen):
            LOG.info("Game mode: skipped snap for fullscreen window hwnd=%s", hwnd)
            self.blocked.emit("Game mode: fullscreen window not moved")
            return

        # Get work area for normal snapping sizing
        mon = get_monitor_rect(hwnd, use_work_area=True)
        if not mon:
            return
        left_edge, top_edge, right_edge, bottom_edge = [int(x) for x in mon]
        monitor_width = int(right_edge - left_edge)
        monitor_height = int(bottom_edge - top_edge)

        # If not a game but is fullscreen, exit fullscreen first (existing behavior for apps)
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if full_screen:
            _exit_fullscreen(hwnd)
            rc = refresh_rect()

        is_resizable = bool(style & win32con.WS_SIZEBOX)

        if is_resizable:
            placement = win32gui.GetWindowPlacement(hwnd)
            if placement[1] == win32con.SW_MAXIMIZE:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            w = int(monitor_width * int(self.settings.width_pct) // 100)
            h = int(monitor_height * int(self.settings.height_pct) // 100)
            x = left_edge + ((monitor_width - w) // 2)
            y = top_edge + ((monitor_height - h) // 2)
            if (rc.right - rc.left, rc.bottom - rc.top, rc.left, rc.top) != (
                w,
                h,
                x,
                y,
            ):
                USER32.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)
        else:
            w = max(660, int(monitor_width * 0.35))
            h = max(260, int(monitor_height * 0.25))
            x = left_edge + ((monitor_width - w) // 2)
            y = top_edge + ((monitor_height - h) // 2)
            USER32.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)

    def _restore(self, hwnd: int):
        from PySide6 import QtWidgets

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
        orig = self._orig_sizes.pop(hwnd, None)
        if not orig:
            return
        was_maximized = orig.get("maximized", False) if isinstance(orig, dict) else False
        rect = orig["rect"] if isinstance(orig, dict) else orig
        mon = get_monitor_rect(hwnd)
        if not mon:
            return
        left_edge, top_edge, right_edge, bottom_edge = mon
        monitor_width, monitor_height = (
            right_edge - left_edge,
            bottom_edge - top_edge,
        )

        rc = wintypes.RECT()
        USER32.GetWindowRect(hwnd, ctypes.byref(rc))
        win_left, win_top, win_right, win_bottom = rc.left, rc.top, rc.right, rc.bottom
        if (
            win_left <= left_edge
            and win_top <= top_edge
            and win_right >= right_edge
            and win_bottom >= bottom_edge
        ):
            return

        if was_maximized:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        else:
            left, top, width, height = rect
            x = left_edge + ((monitor_width - width) // 2)
            y = top_edge + ((monitor_height - height) // 2)
            USER32.MoveWindow(hwnd, x, y, width, height, True)


# ------------------------------------------------------------------------------
# Main window with tray icon
# ------------------------------------------------------------------------------


class MainWindow(QtWidgets.QMainWindow):
    """Main application window with tray icon and QWebEngineView frontend.

    UI rendered by React frontend in QWebEngineView. VireloBridge
    mediates all settings/theme/capture/snap communication. Window is resizable
    via WM_NCHITTEST (min 860x600, default 1000x620).
    """

    key_captured = QtCore.Signal(str)
    snap_key_status = QtCore.Signal(str, int)

    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self.settings.snap_key = str(self.settings.snap_key)
        self.settings.restore_key = str(self.settings.restore_key)
        self.settings.enable_snap = bool(self.settings.enable_snap)
        self.settings.snap_presses = normalize_snap_presses(self.settings.snap_presses)
        self.settings.snap_interval = int(self.settings.snap_interval)
        self.settings.width_pct = int(self.settings.width_pct)
        self.settings.height_pct = int(self.settings.height_pct)
        self.settings.ex_auto_size = bool(getattr(self.settings, "ex_auto_size", False))
        self.settings.game_mode_enabled = bool(self.settings.game_mode_enabled)
        self.settings.run_at_startup = bool(self.settings.run_at_startup)
        self.settings.theme = normalize_theme_mode(str(self.settings.theme), DEFAULTS["theme"])

        self._capture_guard = CaptureGuard()
        self._capture_thread = None
        self._capture_worker = None
        self._capture_target = None
        self._explorer_thread = None
        self._explorer_worker = None

        self.is_first_show = True

        self._theme_mode = self.settings.theme
        self._theme_state = self.settings.theme

        self._theme_timer = QtCore.QTimer(self)
        self._theme_timer.setInterval(2000)
        self._theme_timer.timeout.connect(self._sync_system_theme)

        self.setWindowTitle(APP_TITLE)
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            icon = QtGui.QIcon(icon_path)
        else:
            icon = QtGui.QIcon.fromTheme("applications-system")
        self.setWindowIcon(icon)
        QtWidgets.QApplication.setWindowIcon(icon)

        # Frameless + resizable window.
        self.setWindowFlag(QtCore.Qt.WindowType.FramelessWindowHint, True)
        self.setMinimumSize(860, 600)
        self.resize(1000, 620)

        self.tray_icon = QtWidgets.QSystemTrayIcon(icon, self)
        self.tray_icon.setToolTip("Virelo")
        menu = QtWidgets.QMenu(self)
        open_act = menu.addAction("Open")
        open_act.triggered.connect(self._restore_window)

        self.minimize_to_tray_on_exit = True
        self.action_minimize_on_exit = menu.addAction("Minimize to Tray")
        self.action_minimize_on_exit.setCheckable(True)
        self.action_minimize_on_exit.setChecked(self.minimize_to_tray_on_exit)
        self.action_minimize_on_exit.triggered.connect(self._toggle_minimize_on_exit)

        self.action_run_at_startup = menu.addAction("Run at Startup")
        self.action_run_at_startup.setCheckable(True)
        self.action_run_at_startup.setChecked(bool(self.settings.run_at_startup))
        self.action_run_at_startup.triggered.connect(self._toggle_run_at_startup)

        exit_act = menu.addAction("Quit")
        exit_act.triggered.connect(self._really_quit)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

        self.key_captured.connect(self.on_key_captured)

        # snap_enabled used by business logic (ShiftSnapRestore, _test_snap)
        self.snap_enabled = bool(self.settings.enable_snap)

        # --- Bridge + WebView ---
        self._settings_state = SettingsState(self.settings)
        self._snap_service = SnapService(None)  # shift_mgr set after construction
        self._bridge = VireloBridge(self._settings_state, self._snap_service, parent=self)
        self._bridge.set_main_window(self)
        self._bridge.set_capture_guard(self._capture_guard)

        self.webview = VireloWebView(self._bridge, parent=self)

        # Central widget is just the webview -- React handles all UI
        self.setCentralWidget(self.webview)

        # Route snap_key_status signal to bridge
        self.snap_key_status.connect(self._bridge.snap_status.emit)

        # Shortcuts
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+T"), self, activated=self._toggle_theme)
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+Enter"), self, activated=self._test_snap)
        QtGui.QShortcut(QtGui.QKeySequence("F1"), self, activated=self._show_help)

        # Managers.
        self.shift_mgr = ShiftSnapRestore(self.settings)
        self.shift_mgr.triggered.connect(self.shift_mgr.perform)
        self.shift_mgr.blocked.connect(lambda message: self.snap_key_status.emit(message, 3000))

        # Wire snap_service to shift_mgr
        self._snap_service.set_manager(self.shift_mgr)

        # Thread management.
        self._update_explorer_enabled_state()

        self._apply_theme_mode(self._theme_mode)

    # ------------------------------------------------------------------
    # Tray behavior
    # ------------------------------------------------------------------

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.Type.WindowStateChange:
            if self.isMinimized() and self.minimize_to_tray_on_exit:
                QtCore.QTimer.singleShot(0, self.hide)
        super().changeEvent(event)

    def closeEvent(self, event):
        if self.minimize_to_tray_on_exit:
            event.ignore()
            self.hide()
        else:
            self._stop_background_threads()
            self.shift_mgr.cleanup()
            QtWidgets.QApplication.quit()

    def _on_tray_activated(self, reason):
        if reason in (
            QtWidgets.QSystemTrayIcon.ActivationReason.Trigger,
            QtWidgets.QSystemTrayIcon.ActivationReason.DoubleClick,
        ):
            self._restore_window()

    def _restore_window(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def _really_quit(self):
        self._stop_background_threads()
        self.shift_mgr.cleanup()
        QtWidgets.QApplication.quit()

    def _stop_background_threads(self):
        self._stop_capture_worker()
        self._stop_explorer_worker()
        self._stop_theme_sync()

    # ------------------------------------------------------------------
    # Key capture (preserved -- uses bridge signals for status updates)
    # ------------------------------------------------------------------

    @QtCore.Slot(str)
    def on_key_captured(self, key: str):
        self.settings.snap_key = key
        self._bridge.snap_status.emit(f"Snap key set to {key.upper()}.", 3000)

    def _start_key_capture(self):
        self._begin_key_capture("snap", "Press desired snap key... (Esc to cancel)")

    def _start_restore_key_capture(self):
        self._begin_key_capture("restore", "Press desired restore key... (Esc to cancel)")

    def _begin_key_capture(self, target: str, message: str):
        if not self._capture_guard.try_start():
            self._bridge.snap_status.emit("Key capture already in progress.", 2000)
            return
        self._capture_target = target
        self._bridge.snap_status.emit(message, 0)
        self._bridge.capture_status.emit("capturing")

        self._capture_thread = QtCore.QThread(self)
        self._capture_worker = KeyCaptureWorker()
        self._capture_worker.moveToThread(self._capture_thread)
        self._capture_thread.started.connect(self._capture_worker.run)
        self._capture_worker.captured.connect(self._on_capture_key)
        self._capture_worker.cancelled.connect(self._on_capture_cancelled)
        self._capture_worker.finished.connect(self._capture_thread.quit)
        self._capture_worker.finished.connect(self._capture_worker.deleteLater)
        self._capture_thread.finished.connect(self._capture_thread.deleteLater)
        self._capture_thread.finished.connect(self._on_capture_finished)
        self._capture_thread.start()

    def _on_capture_key(self, key: str):
        key_str = str(key).lower()
        if self._capture_target == "restore":
            self.settings.restore_key = key_str
            if hasattr(self, "shift_mgr"):
                self.shift_mgr.update_restore_key(key_str)
            self._bridge.capture_status.emit("done")
            self._bridge.snap_status.emit(f"Restore key set to {key_str.upper()}.", 3000)
            self._bridge.settings_changed.emit(self._settings_state.get_json())
        else:
            if hasattr(self, "shift_mgr"):
                self.shift_mgr.update_binding(key_str)
            self.key_captured.emit(key_str)
            self._bridge.capture_status.emit("done")
            self._bridge.settings_changed.emit(self._settings_state.get_json())

    def _on_capture_cancelled(self, reason: str):
        message = "Key capture timed out." if reason == "timeout" else "Key capture cancelled."
        self._bridge.capture_status.emit("cancelled" if reason != "timeout" else "timeout")
        self._bridge.snap_status.emit(message, 2000)

    def _on_capture_finished(self):
        self._capture_guard.finish()
        self._capture_thread = None
        self._capture_worker = None
        self._capture_target = None

    def _stop_capture_worker(self):
        worker = getattr(self, "_capture_worker", None)
        thread = getattr(self, "_capture_thread", None)
        if worker is not None:
            try:
                worker.stop()
            except Exception:
                pass
        if thread is not None:
            thread.quit()
            thread.wait(2000)
        self._capture_guard.finish()
        self._capture_thread = None
        self._capture_worker = None
        self._capture_target = None

    # ------------------------------------------------------------------
    # Business logic actions (preserved)
    # ------------------------------------------------------------------

    def _test_snap(self):
        try:
            self.shift_mgr.perform(False)
            self._bridge.snap_status.emit("Snap test applied to the active window.", 2000)
        except Exception:
            self._bridge.snap_status.emit("Could not snap the active window.", 2000)

    def _reset_defaults(self):
        """Reset all settings to defaults (called from bridge.reset_defaults)."""
        defaults = DEFAULTS.copy()
        for key, val in defaults.items():
            setattr(self.settings, key, val)
        self.settings.save()
        self.snap_enabled = bool(defaults["enable_snap"])
        self._update_snap_enabled_state()
        self._update_explorer_autosize_thread()
        self._apply_theme_mode(defaults["theme"])
        if hasattr(self, "shift_mgr"):
            self.shift_mgr.update_binding(defaults["snap_key"])
            self.shift_mgr.update_restore_key(defaults["restore_key"])
            self.shift_mgr.update_press_limit(defaults["snap_presses"])
        self._bridge.settings_changed.emit(self._settings_state.get_json())
        self._bridge.snap_status.emit("Defaults loaded.", 3000)

    def _show_help(self):
        QtWidgets.QMessageBox.information(
            self,
            "Virelo – Help",
            "• Press Count & Interval: how many times and how fast to press the Snap Key.\n"
            "• Hold the Restore Key while pressing to restore original window size.\n"
            "• Width/Height: snapped window size as % of the current monitor.\n"
            "• Explorer Auto-Size: auto-fit columns on folder change (Details view).\n"
            "• Game Mode: when enabled, fullscreen windows (typically games) are skipped.\n\n"
            "Shortcuts:\n"
            "  Ctrl+S = Save Settings,  Ctrl+T = Toggle Theme,\n"
            "  Ctrl+Enter = Test Snap,  F1 = Help",
        )

    # ------------------------------------------------------------------
    # Enable/disable state (delegates to pages)
    # ------------------------------------------------------------------

    def _update_snap_enabled_state(self):
        pass  # React reads snap_enabled from settings via bridge

    def _update_explorer_enabled_state(self):
        self._update_explorer_autosize_thread()

    def _update_explorer_autosize_thread(self, *args):
        """Start/stop the Explorer autosize background thread."""
        LOG.info("_update_explorer_autosize_thread: called")
        app = QtWidgets.QApplication.instance()
        pushed_cursor = False
        if app is not None:
            QtGui.QGuiApplication.setOverrideCursor(QtGui.QCursor(QtCore.Qt.CursorShape.WaitCursor))
            pushed_cursor = True
        try:
            group_enabled = bool(self.settings.ex_auto_size)
            LOG.info("_update_explorer_autosize_thread: group_enabled=%s", group_enabled)
            if not group_enabled:
                LOG.info("Explorer autosize: stopping (disabled or unchecked).")
                self._stop_explorer_worker()
                return

            if self._explorer_thread and self._explorer_thread.isRunning():
                LOG.info("Explorer autosize: worker already running.")
                return

            # Enable debug logging for autosize troubleshooting
            LOG.setLevel(logging.DEBUG)
            LOG.info("Explorer autosize: enabling DEBUG logging for troubleshooting")
            LOG.info("Explorer autosize: log file is at %s", getattr(LOG, "log_path", "unknown"))

            self._explorer_thread = QtCore.QThread(self)
            # Tab-aware autosize with debounce, settle detection,
            # rate limiting, and circuit breakers
            # Schedule: debounce 50ms, then retries at 100ms, 250ms, 500ms, 1s
            self._explorer_worker = ExplorerAutosizeWorker(
                _autosize_explorer_columns_quick,
                _autosize_explorer_columns_full,
                _is_window_interactive,
                schedule=(0.05, 0.1, 0.25, 0.5, 1.0),  # Debounce + retry schedule
            )
            self._explorer_worker.moveToThread(self._explorer_thread)
            self._explorer_thread.started.connect(self._explorer_worker.run)
            self._explorer_worker.finished.connect(self._explorer_thread.quit)
            self._explorer_worker.finished.connect(self._explorer_worker.deleteLater)
            self._explorer_thread.finished.connect(self._explorer_thread.deleteLater)
            self._explorer_thread.finished.connect(self._on_explorer_finished)
            self._explorer_thread.start()
            LOG.info(
                "Explorer autosize: worker started with tab-aware engine, "
                "schedule=(0.05, 0.1, 0.25, 0.5, 1.0)"
            )
        finally:
            if pushed_cursor:
                QtGui.QGuiApplication.restoreOverrideCursor()

    def _stop_explorer_worker(self):
        worker = getattr(self, "_explorer_worker", None)
        thread = getattr(self, "_explorer_thread", None)
        if worker is not None:
            try:
                worker.stop()
            except Exception:
                pass
            # Give the worker time to see the stop flag before we wait on the thread
            # This helps avoid COM calls during shutdown
            import time

            time.sleep(0.05)
        if thread is not None:
            thread.quit()
            # Use longer timeout to allow COM cleanup
            if not thread.wait(3000):
                LOG.warning("Explorer autosize: thread did not stop in time")
        self._explorer_worker = None
        self._explorer_thread = None
        LOG.info("Explorer autosize: worker stopped.")

    def _on_explorer_finished(self):
        self._explorer_worker = None
        self._explorer_thread = None

    def showEvent(self, event):
        super().showEvent(event)
        self._update_snap_enabled_state()
        self._update_explorer_enabled_state()
        if self.is_first_show:
            self.is_first_show = False
            self.repaint()
            QtWidgets.QApplication.processEvents()
            self.center_on_screen()

    def _toggle_minimize_on_exit(self):
        self.minimize_to_tray_on_exit = not self.minimize_to_tray_on_exit
        self.action_minimize_on_exit.setChecked(self.minimize_to_tray_on_exit)

    def _toggle_run_at_startup(self):
        try:
            if self.action_run_at_startup.isChecked():
                create_startup_shortcut()
                self.settings.run_at_startup = True
            else:
                remove_startup_shortcut()
                self.settings.run_at_startup = False
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Error", f"Failed to modify startup shortcut:\n{e}")
            self.action_run_at_startup.setChecked(False)
        self.settings.save()

    def _toggle_theme(self):
        new_mode = toggle_theme_mode(self._theme_mode, get_windows_theme())
        self._apply_theme_mode(new_mode)

    def _apply_theme_mode(self, mode: str):
        self._theme_mode = normalize_theme_mode(mode, DEFAULTS["theme"])
        self.settings.theme = self._theme_mode
        if self._theme_mode == "system":
            self._start_theme_sync()
        else:
            self._stop_theme_sync()
            self.set_theme(self._theme_mode)

    def _start_theme_sync(self):
        if not self._theme_timer.isActive():
            self._theme_timer.start()
        self._sync_system_theme()

    def _stop_theme_sync(self):
        if self._theme_timer.isActive():
            self._theme_timer.stop()

    def _sync_system_theme(self):
        if self._theme_mode != "system":
            return
        effective = resolve_theme("system", get_windows_theme())
        if effective != self._theme_state:
            self.set_theme(effective)

    def set_theme(self, theme: str):
        self._theme_state = theme
        self._bridge.theme_applied.emit(theme)

    def center_on_screen(self):
        cursor_pos = QtGui.QCursor.pos()
        screen = (
            QtWidgets.QApplication.screenAt(cursor_pos) or QtWidgets.QApplication.primaryScreen()
        )
        g = screen.availableGeometry()
        w = self.size()
        x = g.x() + (g.width() - w.width()) // 2
        y = g.y() + (g.height() - w.height()) // 2
        self.move(int(x), int(y))

    # ------------------------------------------------------------------
    # Resizable window via WM_NCHITTEST
    # ------------------------------------------------------------------

    def nativeEvent(self, event_type, message):
        if event_type == b"windows_generic_MSG":
            msg = ctypes.wintypes.MSG.from_address(int(message))
            if msg.message == 0x0084:  # WM_NCHITTEST
                x = msg.lParam & 0xFFFF
                y = (msg.lParam >> 16) & 0xFFFF
                # Convert screen coords to window coords
                pos = self.mapFromGlobal(QtCore.QPoint(x, y))
                rect = self.rect()
                BORDER = 4  # 4px grab zone
                result = 0
                # Edges and corners
                if pos.x() <= BORDER:
                    if pos.y() <= BORDER:
                        result = 13  # HTTOPLEFT
                    elif pos.y() >= rect.height() - BORDER:
                        result = 16  # HTBOTTOMLEFT
                    else:
                        result = 10  # HTLEFT
                elif pos.x() >= rect.width() - BORDER:
                    if pos.y() <= BORDER:
                        result = 14  # HTTOPRIGHT
                    elif pos.y() >= rect.height() - BORDER:
                        result = 17  # HTBOTTOMRIGHT
                    else:
                        result = 11  # HTRIGHT
                elif pos.y() <= BORDER:
                    result = 12  # HTTOP
                elif pos.y() >= rect.height() - BORDER:
                    result = 15  # HTBOTTOM
                if result:
                    return True, result
        return super().nativeEvent(event_type, message)


# ------------------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------------------


def main():
    if not is_admin():
        script = os.path.abspath(sys.argv[0])
        params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
        exe = sys.executable
        if exe.lower().endswith("python.exe"):
            candidate = exe.replace("python.exe", "pythonw.exe")
            if os.path.exists(candidate):
                exe = candidate
        hinst = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", exe, f'"{script}" {params}', None, 1
        )
        if int(hinst) <= 32:
            sys.stderr.write("Elevation failed.\n")
            LOG.error("Elevation failed. ShellExecuteW returned %s.", hinst)
            return
        return  # Elevated child will run the app.

    _enable_dpi_awareness()
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass

    QtCore.QCoreApplication.setOrganizationName(ORGANIZATION)
    QtCore.QCoreApplication.setApplicationName(APP_TITLE)

    mutex = win32event.CreateMutex(None, False, f"Global\\{APP_NAME}_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        return

    app = QtWidgets.QApplication(sys.argv)
    # Safer to set after QApplication exists:
    QtWidgets.QApplication.setQuitOnLastWindowClosed(False)

    if not QtWidgets.QSystemTrayIcon.isSystemTrayAvailable():
        QtWidgets.QMessageBox.critical(None, "Error", "No system tray is available. Exiting.")
        return
    win = MainWindow()

    app.aboutToQuit.connect(lambda: (win._stop_background_threads(), win.shift_mgr.cleanup()))
    atexit.register(lambda: (win._stop_background_threads(), win.shift_mgr.cleanup()))

    win._singleton_mutex = mutex
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
