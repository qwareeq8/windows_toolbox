"""Safe integration checks against the real Windows and Qt dependencies."""

import json
import sys
import uuid
from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import MagicMock

import pytest

pytestmark = [
    pytest.mark.requires_qt,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows integration test."),
]


def test_native_runtime_modules_are_importable():
    """The Windows test environment must load real native packages."""
    import comtypes
    import keyboard
    import PySide6
    import win32gui

    assert PySide6.__version__
    assert callable(win32gui.IsWindow)
    assert callable(keyboard.on_press_key)
    assert comtypes.GUID


def test_settings_round_trip_through_real_qsettings_ini(tmp_path, monkeypatch):
    """Settings persist through a real QSettings backend without touching HKCU."""
    from PySide6 import QtCore

    from virelo.settings.persistence import Settings

    original_qsettings = QtCore.QSettings
    settings_path = tmp_path / "settings.ini"

    class SettingsFactory:
        """Create file-backed QSettings objects for the integration test."""

        Status = original_qsettings.Status

        def __call__(self, organization, application):
            del organization, application
            return original_qsettings(
                str(settings_path),
                original_qsettings.Format.IniFormat,
            )

    monkeypatch.setattr(QtCore, "QSettings", SettingsFactory())
    first = Settings()
    first.width_pct = 64
    first.accent = "teal"
    first.save()

    second = Settings()
    assert second.width_pct == 64
    assert second.accent == "teal"


def test_real_qobject_bridge_reports_folder_task_lifecycle(settings_state, monkeypatch):
    """The real Qt bridge reports start and terminal states around background work."""
    from PySide6 import QtCore

    from virelo.bridge.bridge import VireloBridge
    from virelo.services import explorer_views
    from virelo.services.snap import SnapService

    monkeypatch.setattr(
        explorer_views,
        "apply_details_default",
        lambda: {
            "ok": True,
            "data": {"backup": r"C:\backup", "restarted": False},
        },
    )
    bridge = VireloBridge(settings_state, SnapService(None))
    lifecycle = []
    bridge.views_task_changed.connect(lambda payload: lifecycle.append(json.loads(payload)))

    started = json.loads(bridge.apply_details_view())
    assert bridge.wait_for_view_task() is True
    app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])
    app.processEvents()

    assert started["ok"] is True
    assert [event["state"] for event in lifecycle] == ["started", "succeeded"]
    assert lifecycle[-1]["data"]["backup"] == r"C:\backup"


def test_capture_shutdown_retains_references_when_qthread_is_still_live():
    """A timed-out capture stop must not abandon a running QThread."""
    from virelo.app.window import MainWindow

    worker = SimpleNamespace(stop=MagicMock())
    thread = SimpleNamespace(quit=MagicMock(), wait=MagicMock(return_value=False))
    guard = SimpleNamespace(finish=MagicMock())
    owner = SimpleNamespace(
        _capture_worker=worker,
        _capture_thread=thread,
        _capture_guard=guard,
        _capture_target="snap",
    )

    assert MainWindow._stop_capture_worker(cast(Any, owner), timeout_ms=10) is False
    assert owner._capture_worker is worker
    assert owner._capture_thread is thread
    guard.finish.assert_not_called()


def test_webview_dev_navigation_requires_the_exact_vite_origin(monkeypatch):
    """Development navigation does not accept arbitrary localhost origins."""
    from PySide6.QtCore import QUrl

    from virelo.app.webview import _is_dev_server_url

    monkeypatch.setenv("VIRELO_DEV", "1")
    assert _is_dev_server_url(QUrl("http://localhost:5173/settings")) is True
    assert _is_dev_server_url(QUrl("http://localhost:5174/settings")) is False
    assert _is_dev_server_url(QUrl("https://localhost:5173/settings")) is False
    assert _is_dev_server_url(QUrl("http://127.0.0.1:5173/settings")) is False


def test_webview_local_navigation_stays_inside_frontend_root(tmp_path, monkeypatch):
    """Release navigation cannot use the elevated webview to open other files."""
    from PySide6.QtCore import QUrl

    from virelo.app import webview

    frontend = tmp_path / "frontend" / "dist"
    frontend.mkdir(parents=True)
    inside = frontend / "index.html"
    outside = tmp_path / "private.txt"
    inside.write_text("Virelo", encoding="utf-8")
    outside.write_text("Private", encoding="utf-8")
    monkeypatch.setattr(
        webview,
        "resource_path",
        lambda relative: str(tmp_path / relative),
    )

    assert webview._is_allowed_frontend_file(QUrl.fromLocalFile(str(inside))) is True
    assert webview._is_allowed_frontend_file(QUrl.fromLocalFile(str(outside))) is False


def test_registry_helpers_copy_verify_and_delete_isolated_test_keys():
    """Registry helpers work against an isolated HKCU subtree and clean it up."""
    import winreg

    from virelo.services.explorer_views import (
        RegValue,
        _copy_key_tree,
        _delete_key_tree,
        _restore_key_tree,
        _snapshot_key_tree,
        _write_values,
    )

    test_root = rf"Software\Virelo\Tests\{uuid.uuid4()}"
    source = test_root + r"\Source"
    destination = test_root + r"\Destination"
    snapshot_destination = test_root + r"\SnapshotDestination"
    access = winreg.KEY_READ | winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, source + r"\Child", 0, access) as key:
            winreg.SetValueEx(key, "Original", 0, winreg.REG_SZ, "value")
            winreg.SetValueEx(key, "Binary", 0, winreg.REG_BINARY, b"\x00\x01\xff")
            winreg.SetValueEx(key, "Multi", 0, winreg.REG_MULTI_SZ, ["one", "two"])
            winreg.SetValueEx(key, "Qword", 0, winreg.REG_QWORD, 2**40)

        _copy_key_tree(
            winreg,
            winreg.HKEY_CURRENT_USER,
            source,
            winreg.HKEY_CURRENT_USER,
            destination,
        )
        _write_values(
            winreg,
            [RegValue(destination + r"\Child", "Verified", "dword", 42)],
        )
        snapshot = _snapshot_key_tree(winreg, winreg.HKEY_CURRENT_USER, source)
        _restore_key_tree(
            winreg,
            winreg.HKEY_CURRENT_USER,
            snapshot_destination,
            snapshot,
        )

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            destination + r"\Child",
            0,
            access,
        ) as key:
            assert winreg.QueryValueEx(key, "Original") == ("value", winreg.REG_SZ)
            assert winreg.QueryValueEx(key, "Verified") == (42, winreg.REG_DWORD)
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            snapshot_destination + r"\Child",
            0,
            access,
        ) as key:
            assert winreg.QueryValueEx(key, "Original") == ("value", winreg.REG_SZ)
            assert winreg.QueryValueEx(key, "Binary") == (b"\x00\x01\xff", winreg.REG_BINARY)
            assert winreg.QueryValueEx(key, "Multi") == (["one", "two"], winreg.REG_MULTI_SZ)
            assert winreg.QueryValueEx(key, "Qword") == (2**40, winreg.REG_QWORD)
    finally:
        _delete_key_tree(winreg, winreg.HKEY_CURRENT_USER, test_root)

    with pytest.raises(FileNotFoundError):
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, test_root)


def test_retained_pywin32_identity_matches_the_same_comtypes_explorer_tab():
    """Cross-library IUnknown matching targets a tab without using its path."""
    import logging

    import pythoncom
    import win32com.client
    import win32gui

    from virelo.services.explorer_columns import find_explorer_tab_by_identity

    pythoncom.CoInitialize()
    try:
        windows = win32com.client.Dispatch("Shell.Application").Windows()
        retained = None
        candidate = None
        for index in range(int(windows.Count)):
            candidate = windows.Item(index)
            if candidate is None:
                continue
            hwnd = int(candidate.HWND or 0)
            if hwnd and win32gui.GetClassName(hwnd) in ("CabinetWClass", "ExploreWClass"):
                retained = candidate
                break
        if retained is None:
            pytest.skip("No File Explorer tab is open for the identity integration check.")

        retained_hwnd = int(retained.HWND)
        identity = retained._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        found = find_explorer_tab_by_identity(
            logging.getLogger("test"),
            retained_hwnd,
            identity,
        )
        found_hwnd = None if found is None else found.hwnd
        del found, identity, retained, candidate, windows
    finally:
        pythoncom.CoUninitialize()

    assert found_hwnd == retained_hwnd
