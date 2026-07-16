"""Focused tests for Explorer tab matching and autosize completion."""

import ctypes
import sys
import types
from unittest.mock import patch, sentinel


def _complete_comtypes_stub() -> None:
    """Complete the lightweight root test stub enough to import COM declarations."""
    comtypes = sys.modules["comtypes"]
    if hasattr(comtypes, "COMMETHOD"):
        return

    class GUID:
        def __init__(self, value):
            self.value = value

        def __str__(self):
            return self.value

    class IUnknown:
        pass

    class COMError(Exception):
        @property
        def hresult(self):
            return self.args[0] if self.args else 0

    client = types.ModuleType("comtypes.client")
    setattr(client, "CreateObject", lambda name: None)
    sys.modules["comtypes.client"] = client
    setattr(comtypes, "client", client)
    setattr(comtypes, "COMMETHOD", lambda *args, **kwargs: (args, kwargs))
    setattr(comtypes, "GUID", GUID)
    setattr(comtypes, "HRESULT", ctypes.c_long)
    setattr(comtypes, "POINTER", ctypes.POINTER)
    setattr(comtypes, "IUnknown", IUnknown)
    setattr(comtypes, "byref", ctypes.byref)
    setattr(comtypes, "cast", ctypes.cast)
    setattr(comtypes, "COMError", COMError)
    setattr(comtypes, "CoInitialize", lambda: None)
    setattr(comtypes, "CoUninitialize", lambda: None)


_complete_comtypes_stub()

from virelo.services.explorer_columns import (  # noqa: E402
    FVM_DETAILS,
    ShellWindow,
    _resolve_shell_window_location,
    autosize_explorer_columns,
    autosize_explorer_columns_detailed,
    find_explorer_tab_by_identity,
    find_explorer_tab_by_path,
)


def _shell_window(path: str, *, tab_id: int = 7) -> ShellWindow:
    return ShellWindow(
        dispatch=object(),
        hwnd=101,
        exe_name="explorer.exe",
        location_url=path,
        tab_id=tab_id,
        view_mode=FVM_DETAILS,
    )


def test_find_tab_by_path_matches_unc_file_url():
    """A Shell file URL should match the equivalent canonical UNC path."""
    candidate = _shell_window("file://Server/Share/Folder%20Name")

    with patch(
        "virelo.services.explorer_columns.iter_explorer_tabs",
        return_value=iter([candidate]),
    ):
        found = find_explorer_tab_by_path(
            __import__("logging").getLogger("test"),
            101,
            r"\\server\share\folder name",
        )

    assert found is candidate


def test_find_tab_by_identity_targets_duplicate_path_after_reordering():
    """IUnknown matching must select the same duplicate-path tab in any order."""
    first = _shell_window(r"C:\shared", tab_id=1)
    target = _shell_window(r"C:\shared", tab_id=2)
    target = ShellWindow(
        dispatch=sentinel.target_dispatch,
        hwnd=target.hwnd,
        exe_name=target.exe_name,
        location_url=target.location_url,
        tab_id=target.tab_id,
        view_mode=target.view_mode,
    )
    logger = __import__("logging").getLogger("test")

    for candidates in ([first, target], [target, first]):
        with (
            patch(
                "virelo.services.explorer_columns.iter_explorer_tabs",
                return_value=iter(candidates),
            ),
            patch(
                "virelo.services.explorer_columns._dispatch_matches_identity",
                side_effect=lambda dispatch, identity: (
                    dispatch is sentinel.target_dispatch and identity is sentinel.identity
                ),
            ),
        ):
            found = find_explorer_tab_by_identity(logger, 101, sentinel.identity)

        assert found is target


def test_modern_autosize_never_falls_back_to_duplicate_path_matching():
    """A retained identity must be the sole selector for modern worker calls."""
    candidate = _shell_window(r"C:\shared", tab_id=2)
    with (
        patch(
            "virelo.services.explorer_columns.find_explorer_tab_by_identity",
            return_value=candidate,
        ) as by_identity,
        patch(
            "virelo.services.explorer_columns.find_explorer_tab_by_path",
            side_effect=AssertionError("path-first dispatch is unsafe for duplicate tabs"),
        ),
        patch("virelo.services.explorer_columns.apply_to_window", return_value=(3, 3, 4)),
    ):
        result = autosize_explorer_columns(
            101,
            target_path=r"C:\shared",
            caller_owns_com=True,
            tab_identity=sentinel.identity,
        )

    assert result == (True, "com")
    by_identity.assert_called_once_with(
        __import__("logging").getLogger("Virelo"),
        101,
        sentinel.identity,
    )


def test_shell_location_falls_back_for_virtual_folder():
    """Virtual folders with no URL should use their document path or display name."""
    item = type("Item", (), {"Path": "This PC"})()
    folder = type("Folder", (), {"Self": item})()
    document = type("Document", (), {"Folder": folder})()
    dispatch = type(
        "Dispatch",
        (),
        {"LocationURL": "", "Document": document, "LocationName": "Computer"},
    )()

    assert _resolve_shell_window_location(dispatch) == "This PC"


def test_simple_autosize_requires_every_visible_column_to_succeed():
    """Partial column success must remain retryable instead of entering dedupe."""
    candidate = _shell_window(r"C:\work")
    with (
        patch(
            "virelo.services.explorer_columns.find_explorer_window_by_hwnd",
            return_value=candidate,
        ),
        patch("virelo.services.explorer_columns.apply_to_window", return_value=(3, 2, 4)),
    ):
        result = autosize_explorer_columns(101, caller_owns_com=True)

    assert result == (False, "none")


def test_detailed_autosize_reports_partial_completion_as_transient():
    """Detailed callers should receive the partial counts and retry guidance."""
    candidate = _shell_window(r"C:\work")
    with (
        patch(
            "virelo.services.explorer_columns.find_active_tab_for_hwnd",
            return_value=candidate,
        ),
        patch("virelo.services.explorer_columns.apply_to_window", return_value=(3, 2, 4)),
    ):
        result = autosize_explorer_columns_detailed(101, caller_owns_com=True)

    assert result.success is False
    assert result.transient_error is True
    assert result.columns_attempted == 3
    assert result.columns_succeeded == 2
    assert result.error_message == "Autosized 2 of 3 visible columns"
