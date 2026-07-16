"""Snap service, MultiPressHotkeyListener, and SnapRestoreController engine.

MultiPressHotkeyListener detects multi-press keyboard patterns and emits a trigger signal.
SnapRestoreController performs window snap and restore operations.
SnapService wraps both so bridge.py can trigger snap actions without depending
on class internals.
"""

import ctypes
import logging
import threading
import time
from collections import deque
from ctypes import wintypes
from typing import cast

import keyboard
import win32con
import win32gui
from PySide6 import QtCore

from virelo.app.config import normalize_snap_presses
from virelo.platform.win32_helpers import (
    USER32,
    _exit_fullscreen,
    _get_window_dwm_rect,
    _is_window_fullscreen,
    _should_skip_snap_for_game,
    get_monitor_rect,
)

LOG = logging.getLogger("Virelo")


def window_border_deltas(
    win_rect: tuple[int, int, int, int],
    visible_rect: tuple[int, int, int, int] | None,
) -> tuple[int, int, int, int]:
    """Return (left, top, right, bottom) invisible-border widths.

    DWM windows extend past their visible frame; centering on the raw window
    rect leaves the window visually off-center. Deltas are clamped to a sane
    range so a bogus DWM answer cannot fling the window off-screen.
    """
    if visible_rect is None:
        return (0, 0, 0, 0)
    wl, wt, wr, wb = win_rect
    vl, vt, vr, vb = visible_rect
    clamp = lambda v: max(0, min(64, v))  # noqa: E731
    return (clamp(vl - wl), clamp(vt - wt), clamp(wr - vr), clamp(wb - vb))


def calculate_snap_position(
    monitor_left: int,
    monitor_top: int,
    monitor_width: int,
    monitor_height: int,
    width_pct: int,
    height_pct: int,
) -> tuple:
    """Calculate snap target position (x, y, w, h) for a resizable window."""
    w = monitor_width * width_pct // 100
    h = monitor_height * height_pct // 100
    x = monitor_left + (monitor_width - w) // 2
    y = monitor_top + (monitor_height - h) // 2
    return (x, y, w, h)


def calculate_centered_window_position(
    monitor_left: int,
    monitor_top: int,
    monitor_width: int,
    monitor_height: int,
    window_width: int,
    window_height: int,
    borders: tuple[int, int, int, int] = (0, 0, 0, 0),
) -> tuple[int, int, int, int]:
    """Center a fixed-size window by its visible frame without resizing it."""
    border_left, border_top, border_right, border_bottom = borders
    visible_width = max(1, window_width - border_left - border_right)
    visible_height = max(1, window_height - border_top - border_bottom)
    x = monitor_left + ((monitor_width - visible_width) // 2) - border_left
    y = monitor_top + ((monitor_height - visible_height) // 2) - border_top
    return (x, y, window_width, window_height)


class MultiPressHotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""

    triggered = QtCore.Signal(bool)

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self._press_times: deque[float] = deque(
            maxlen=normalize_snap_presses(self.settings.snap_presses)
        )
        self._press_lock = threading.Lock()
        self._held = False
        self._last_press_event = 0.0
        self.current_key = str(settings.snap_key)
        self.restore_key = str(getattr(settings, "restore_key", "ctrl"))
        self._press_hook = None
        self._release_hook = None
        try:
            self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
            self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)
        except Exception:
            if self._press_hook is not None:
                try:
                    keyboard.unhook(self._press_hook)
                except Exception:
                    LOG.exception("Rolling back the initial snap-key hook failed.")
            self._press_hook = None
            self._release_hook = None
            self.settings.enable_snap = False
            LOG.exception(
                "Installing the global snap-key hooks failed; snapping was disabled for this run."
            )

    def cleanup(self):
        for hook in (self._press_hook, self._release_hook):
            if hook is None:
                continue
            try:
                keyboard.unhook(hook)
            except Exception:
                pass
        self._press_hook = None
        self._release_hook = None
        self._held = False

    def update_binding(self, new_key: str) -> bool:
        # Install the new hooks BEFORE removing the old ones and roll back on
        # failure, so a malformed key name cannot leave snapping dead with no
        # working hook installed.
        new_press = None
        try:
            new_press = keyboard.on_press_key(new_key, self._on_press)
            new_release = keyboard.on_release_key(new_key, self._on_release)
        except Exception:
            if new_press is not None:
                try:
                    keyboard.unhook(new_press)
                except Exception:
                    LOG.exception("Rolling back the partial snap-key binding failed.")
            LOG.exception("Rebinding snap key to %r failed; keeping the current binding", new_key)
            return False
        for hook in (self._press_hook, self._release_hook):
            if hook is None:
                continue
            try:
                keyboard.unhook(hook)
            except Exception:
                pass
        # The old release hook is gone; a key physically held through the swap
        # would otherwise leave _held stuck True forever.
        self._held = False
        self.current_key = str(new_key).strip()
        self._press_hook = new_press
        self._release_hook = new_release
        with self._press_lock:
            self._press_times = deque(
                self._press_times,
                maxlen=normalize_snap_presses(self.settings.snap_presses),
            )
        return True

    def update_restore_key(self, new_key: str):
        self.restore_key = new_key

    def update_press_limit(self, new_limit: int):
        with self._press_lock:
            self._press_times = deque(
                self._press_times,
                maxlen=normalize_snap_presses(new_limit),
            )

    def _on_press(self, event):
        if not self.settings.enable_snap:
            return
        now = time.monotonic()
        # Recover from a lost release event (focus steal, UAC prompt, session
        # switch). A genuinely held key produces auto-repeat press events well
        # under a second apart, so a long-silent "held" state is stale.
        if self._held and now - self._last_press_event > 1.0:
            self._held = False
        self._last_press_event = now
        if not self._held:
            self._held = True
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


# ------------------------------------------------------------------------------
# SHIFT triple-press snap and restore
# ------------------------------------------------------------------------------


class SnapRestoreController(QtCore.QObject):
    """Performs window snap and restore operations."""

    blocked = QtCore.Signal(str)

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        # Original geometry is captured at snap time (in _snap), not at
        # startup. Pre-populating here would make restore return a window to
        # wherever it happened to be when Virelo launched, not to where it was
        # right before the user snapped it.
        self._orig_sizes: dict[int, dict[str, tuple[int, int, int, int] | bool]] = {}

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
                        "borders": window_border_deltas(
                            (rc.left, rc.top, rc.right, rc.bottom),
                            _get_window_dwm_rect(hwnd),
                        ),
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
                            "borders": window_border_deltas(
                                (rc.left, rc.top, rc.right, rc.bottom),
                                _get_window_dwm_rect(hwnd),
                            ),
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

    # Shell windows that must never be moved or resized.
    _SHELL_CLASSES = frozenset(
        {
            "Shell_TrayWnd",
            "Shell_SecondaryTrayWnd",
            "Progman",
            "WorkerW",
            "NotifyIconOverflowWindow",
        }
    )

    @QtCore.Slot(bool)
    def perform(self, restore: bool) -> bool:
        """Snap or restore the foreground window. Returns True if it acted."""
        self._prune_closed_windows()
        hwnd = USER32.GetForegroundWindow()
        if not hwnd:
            return False
        try:
            if win32gui.GetClassName(hwnd) in self._SHELL_CLASSES:
                LOG.debug("Snap: skipping shell window hwnd=%s", hwnd)
                return False
        except Exception:
            pass
        try:
            if restore:
                return bool(self._restore(hwnd))
            return bool(self._snap(hwnd))
        except Exception as e:
            LOG.exception("SnapRestoreController.perform failed.", exc_info=e)
            return False

    def _snap(self, hwnd: int) -> bool:
        from PySide6 import QtWidgets

        for widget in QtWidgets.QApplication.topLevelWidgets():
            if int(widget.winId()) == hwnd:
                LOG.debug("Snap: skipping Virelo's own window hwnd=%s", hwnd)
                return False  # Skip entirely per SNAP-03

        def refresh_rect():
            rect = wintypes.RECT()
            USER32.GetWindowRect(hwnd, ctypes.byref(rect))
            return rect

        rc = refresh_rect()

        # Get full monitor bounds for accurate fullscreen detection
        mon_full = get_monitor_rect(hwnd, use_work_area=False)
        if not mon_full:
            return False

        # Check if window is fullscreen using full monitor bounds. Let the
        # helper fetch the DWM extended frame itself: the raw window rect
        # includes invisible borders that defeat the tolerance check.
        full_screen = _is_window_fullscreen(hwnd, monitor_rect=mon_full)

        # Skip snapping if game mode enabled and window is fullscreen borderless
        if _should_skip_snap_for_game(hwnd, self.settings, full_screen):
            LOG.info("Game mode: skipped snap for fullscreen window hwnd=%s", hwnd)
            self.blocked.emit("Game mode: fullscreen window not moved")
            return False

        # Get work area for normal snapping sizing
        mon = get_monitor_rect(hwnd, use_work_area=True)
        if not mon:
            return False
        left_edge, top_edge, right_edge, bottom_edge = [int(x) for x in mon]
        monitor_width = int(right_edge - left_edge)
        monitor_height = int(bottom_edge - top_edge)

        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        window_width = int(rc.right - rc.left)
        window_height = int(rc.bottom - rc.top)
        if monitor_width <= 0 or monitor_height <= 0 or window_width <= 0 or window_height <= 0:
            return False

        # Capture geometry only after every eligibility and monitor check has
        # passed. A skipped snap must not create a restore entry for a window
        # that Virelo never moved.
        placement = win32gui.GetWindowPlacement(hwnd)
        was_maximized = placement[1] == win32con.SW_MAXIMIZE
        original_borders = window_border_deltas(
            (rc.left, rc.top, rc.right, rc.bottom), _get_window_dwm_rect(hwnd)
        )
        if hwnd not in self._orig_sizes:
            self._orig_sizes[hwnd] = {
                "rect": (rc.left, rc.top, window_width, window_height),
                "maximized": was_maximized,
                "borders": original_borders,
            }

        # If not a game but is fullscreen, exit fullscreen first (existing behavior for apps)
        if full_screen:
            _exit_fullscreen(hwnd)
            rc = refresh_rect()

        is_resizable = bool(style & win32con.WS_SIZEBOX)

        if is_resizable:
            if was_maximized:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                rc = refresh_rect()
            # Center the VISIBLE frame: size the visible box to the requested
            # percentages, then widen the window rect by the invisible DWM
            # borders and shift left/up so the visible frame lands centered.
            visible_w = int(monitor_width * int(self.settings.width_pct) // 100)
            visible_h = int(monitor_height * int(self.settings.height_pct) // 100)
            border_l, border_t, border_r, border_b = window_border_deltas(
                (rc.left, rc.top, rc.right, rc.bottom), _get_window_dwm_rect(hwnd)
            )
            w = visible_w + border_l + border_r
            h = visible_h + border_t + border_b
            x = left_edge + ((monitor_width - visible_w) // 2) - border_l
            y = top_edge + ((monitor_height - visible_h) // 2) - border_t
            if (rc.right - rc.left, rc.bottom - rc.top, rc.left, rc.top) == (w, h, x, y):
                return True  # Already at the target; nothing to do.
            return bool(USER32.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True))
        else:
            # Fixed-size windows such as Google Drive and LightBulb can render
            # incorrectly when MoveWindow is asked to resize them. Preserve
            # the raw window dimensions and center the DWM-visible frame.
            x, y, w, h = calculate_centered_window_position(
                left_edge,
                top_edge,
                monitor_width,
                monitor_height,
                int(rc.right - rc.left),
                int(rc.bottom - rc.top),
                original_borders,
            )
            if (rc.left, rc.top) == (x, y):
                return True
            return bool(USER32.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True))

    def _restore(self, hwnd: int) -> bool:
        from PySide6 import QtWidgets

        for widget in QtWidgets.QApplication.topLevelWidgets():
            if int(widget.winId()) == hwnd:
                LOG.debug("Restore: skipping Virelo's own window hwnd=%s", hwnd)
                return False  # Skip entirely per D-05
        # Peek, do not pop yet: if we cannot resolve a monitor we keep the
        # saved geometry so a later restore can still succeed.
        orig = self._orig_sizes.get(hwnd)
        if not orig:
            return False
        mon = get_monitor_rect(hwnd)
        if not mon:
            return False
        was_maximized = orig.get("maximized", False) if isinstance(orig, dict) else False
        rect = cast(
            tuple[int, int, int, int],
            orig["rect"] if isinstance(orig, dict) else orig,
        )
        left_edge, top_edge, right_edge, bottom_edge = mon

        if was_maximized:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        else:
            left, top, width, height = rect
            saved_borders = cast(
                tuple[int, int, int, int],
                orig.get("borders", (0, 0, 0, 0)),
            )
            border_left, border_top, border_right, border_bottom = saved_borders
            visible_width = max(1, width - border_left - border_right)
            visible_height = max(1, height - border_top - border_bottom)

            # Clamp the visible DWM frame, not the raw window rectangle. The
            # raw rectangle can legitimately extend past the work area by its
            # invisible border, as Google Drive does at the taskbar edge.
            visible_left = left + border_left
            visible_top = top + border_top
            clamped_visible_left = max(left_edge, min(visible_left, right_edge - visible_width))
            clamped_visible_top = max(top_edge, min(visible_top, bottom_edge - visible_height))
            x = clamped_visible_left - border_left
            y = clamped_visible_top - border_top
            if not USER32.MoveWindow(hwnd, int(x), int(y), int(width), int(height), True):
                return False
        # Restore succeeded: forget the saved geometry so a re-snap captures
        # the new pre-snap position.
        self._orig_sizes.pop(hwnd, None)
        return True


class SnapService:
    """Narrow API surface for snap/restore actions."""

    def __init__(self, shift_mgr):
        """Accept a SnapRestoreController instance (or None during early init)."""
        self._mgr = shift_mgr
        self._listener = None

    def set_manager(self, mgr):
        """Set or replace the SnapRestoreController instance."""
        self._mgr = mgr

    def set_listener(self, listener):
        """Set or replace the MultiPressHotkeyListener instance."""
        self._listener = listener

    def test_snap(self) -> dict:
        """Trigger a test snap (same as pressing "Test snap" button)."""
        if self._mgr is None:
            return {"ok": False, "error": "Snap manager not initialized"}
        try:
            acted = self._mgr.perform(False)
            if acted:
                return {"ok": True, "message": "Snap test applied to the active window."}
            return {"ok": False, "error": "No window was snapped (nothing eligible in front)."}
        except Exception as e:
            LOG.exception("test_snap failed")
            return {"ok": False, "error": str(e)}

    def update_binding(self, key: str):
        """Update the snap key binding."""
        if self._listener:
            self._listener.update_binding(key)

    def update_restore_key(self, key: str):
        """Update the restore key binding."""
        if self._listener:
            self._listener.update_restore_key(key)

    def update_press_limit(self, count: int):
        """Update the snap press count."""
        if self._listener:
            self._listener.update_press_limit(count)
