"""Regression tests for fixed-size window centering and restore capture."""

import ctypes
from ctypes import wintypes
from types import SimpleNamespace
from unittest.mock import MagicMock

import PySide6.QtWidgets as QtWidgets

from virelo.app.config import DEFAULTS
from virelo.services import snap


def _settings(**overrides):
    values = {**DEFAULTS, **overrides}
    return SimpleNamespace(**values)


def _controller():
    return snap.SnapRestoreController(_settings())


def _install_empty_qt_window_list(monkeypatch):
    application = MagicMock()
    application.topLevelWidgets.return_value = []
    monkeypatch.setattr(QtWidgets, "QApplication", application, raising=False)


def _user32_with_rect(rect):
    user32 = MagicMock()

    def get_window_rect(_hwnd, rect_pointer):
        output = ctypes.cast(rect_pointer, ctypes.POINTER(wintypes.RECT)).contents
        output.left, output.top, output.right, output.bottom = rect
        return True

    user32.GetWindowRect.side_effect = get_window_rect
    user32.MoveWindow.return_value = True
    return user32


def test_fixed_size_window_centers_visible_frame_without_resizing(monkeypatch):
    """A window without WS_SIZEBOX must keep its dimensions when centered."""
    hwnd = 123
    controller = _controller()
    user32 = _user32_with_rect((100, 100, 900, 700))
    _install_empty_qt_window_list(monkeypatch)
    monkeypatch.setattr(snap, "USER32", user32)
    monkeypatch.setattr(
        snap,
        "get_monitor_rect",
        MagicMock(side_effect=[(0, 0, 1920, 1080), (0, 0, 1920, 1040)]),
    )
    monkeypatch.setattr(snap, "_is_window_fullscreen", MagicMock(return_value=False))
    monkeypatch.setattr(snap, "_should_skip_snap_for_game", MagicMock(return_value=False))
    monkeypatch.setattr(snap, "_get_window_dwm_rect", MagicMock(return_value=(107, 100, 893, 693)))
    monkeypatch.setattr(snap.win32gui, "GetWindowLong", MagicMock(return_value=0))
    monkeypatch.setattr(
        snap.win32gui,
        "GetWindowPlacement",
        MagicMock(return_value=(0, 1, (0, 0), (0, 0), (100, 100, 900, 700))),
    )

    assert controller._snap(hwnd) is True

    # The visible 786 by 593 frame is centered in the work area, while the
    # raw 800 by 600 window size is preserved exactly.
    user32.MoveWindow.assert_called_once_with(hwnd, 560, 223, 800, 600, True)
    assert controller._orig_sizes[hwnd] == {
        "rect": (100, 100, 800, 600),
        "maximized": False,
        "borders": (7, 0, 7, 7),
    }


def test_restore_clamps_visible_frame_and_preserves_raw_geometry(monkeypatch):
    """Restore must allow an invisible border to extend below the work area."""
    hwnd = 321
    controller = _controller()
    controller._orig_sizes[hwnd] = {
        "rect": (1273, 660, 1040, 739),
        "maximized": False,
        "borders": (7, 0, 7, 7),
    }
    user32 = MagicMock()
    user32.MoveWindow.return_value = True
    _install_empty_qt_window_list(monkeypatch)
    monkeypatch.setattr(snap, "USER32", user32)
    monkeypatch.setattr(snap, "get_monitor_rect", MagicMock(return_value=(0, 0, 2560, 1392)))

    assert controller._restore(hwnd) is True

    user32.MoveWindow.assert_called_once_with(hwnd, 1273, 660, 1040, 739, True)
    assert hwnd not in controller._orig_sizes


def test_failed_monitor_lookup_does_not_create_restore_geometry(monkeypatch):
    """A monitor lookup failure must leave restore state untouched."""
    hwnd = 456
    controller = _controller()
    user32 = _user32_with_rect((100, 100, 900, 700))
    _install_empty_qt_window_list(monkeypatch)
    monkeypatch.setattr(snap, "USER32", user32)
    monkeypatch.setattr(snap, "get_monitor_rect", MagicMock(return_value=None))

    assert controller._snap(hwnd) is False
    assert hwnd not in controller._orig_sizes
    user32.MoveWindow.assert_not_called()


def test_failed_work_area_lookup_does_not_create_restore_geometry(monkeypatch):
    """A work-area lookup failure must not record a window as snapped."""
    hwnd = 789
    controller = _controller()
    user32 = _user32_with_rect((100, 100, 900, 700))
    _install_empty_qt_window_list(monkeypatch)
    monkeypatch.setattr(snap, "USER32", user32)
    monkeypatch.setattr(
        snap,
        "get_monitor_rect",
        MagicMock(side_effect=[(0, 0, 1920, 1080), None]),
    )
    monkeypatch.setattr(snap, "_is_window_fullscreen", MagicMock(return_value=False))
    monkeypatch.setattr(snap, "_should_skip_snap_for_game", MagicMock(return_value=False))

    assert controller._snap(hwnd) is False
    assert hwnd not in controller._orig_sizes
    user32.MoveWindow.assert_not_called()


def test_game_mode_skip_does_not_create_restore_geometry(monkeypatch):
    """A fullscreen game-mode skip must not create a restore entry."""
    hwnd = 987
    controller = _controller()
    user32 = _user32_with_rect((0, 0, 1920, 1080))
    _install_empty_qt_window_list(monkeypatch)
    monkeypatch.setattr(snap, "USER32", user32)
    monkeypatch.setattr(snap, "get_monitor_rect", MagicMock(return_value=(0, 0, 1920, 1080)))
    monkeypatch.setattr(snap, "_is_window_fullscreen", MagicMock(return_value=True))
    monkeypatch.setattr(snap, "_should_skip_snap_for_game", MagicMock(return_value=True))

    assert controller._snap(hwnd) is False
    assert hwnd not in controller._orig_sizes
    user32.MoveWindow.assert_not_called()
