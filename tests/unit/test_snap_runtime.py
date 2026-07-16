"""Regression tests for snap restore failure handling."""

from unittest.mock import MagicMock, patch

from virelo.services import snap


def test_failed_restore_retains_geometry_for_retry(mock_settings):
    """A failed MoveWindow must not report success or consume saved geometry."""
    manager = snap.SnapRestoreController.__new__(snap.SnapRestoreController)
    manager.settings = mock_settings
    manager._orig_sizes = {101: {"rect": (100, 100, 800, 600), "maximized": False}}
    application = MagicMock()
    application.topLevelWidgets.return_value = []

    with (
        patch("PySide6.QtWidgets.QApplication", application, create=True),
        patch("virelo.services.snap.USER32") as user32,
        patch("virelo.services.snap.get_monitor_rect", return_value=(0, 0, 1920, 1080)),
    ):
        user32.MoveWindow.return_value = False
        assert manager._restore(101) is False

    assert 101 in manager._orig_sizes
