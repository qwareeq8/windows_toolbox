"""Tests for snap hotkey timing and binding replacement."""

from types import SimpleNamespace
from unittest.mock import MagicMock

from virelo.services import snap


def _settings(**overrides):
    values = {
        "enable_snap": True,
        "snap_presses": 2,
        "snap_interval": 500,
        "snap_key": "shift",
        "restore_key": "ctrl",
    }
    values.update(overrides)
    return SimpleNamespace(**values)


def test_press_detection_uses_monotonic_time(monkeypatch):
    """Two presses within the interval must use monotonic elapsed time."""
    monkeypatch.setattr(
        snap.keyboard, "on_press_key", MagicMock(return_value="old-press"), raising=False
    )
    monkeypatch.setattr(
        snap.keyboard,
        "on_release_key",
        MagicMock(return_value="old-release"),
        raising=False,
    )
    monkeypatch.setattr(snap.keyboard, "is_pressed", MagicMock(return_value=False), raising=False)
    monotonic = MagicMock(side_effect=[10.0, 10.2])
    monkeypatch.setattr(snap.time, "monotonic", monotonic)
    listener = snap.MultiPressHotkeyListener(_settings())
    emissions: list[bool] = []
    signal_emit = listener.triggered.emit
    if hasattr(signal_emit, "reset_mock"):
        signal_emit.reset_mock()
    else:
        listener.triggered.connect(emissions.append)

    listener._on_press(None)
    listener._on_release(None)
    listener._on_press(None)

    assert monotonic.call_count == 2
    if hasattr(signal_emit, "assert_called_once_with"):
        signal_emit.assert_called_once_with(False)
    else:
        assert emissions == [False]


def test_binding_failure_unhooks_partial_new_binding(monkeypatch):
    """A release-hook failure must remove the newly installed press hook."""
    press_hook = MagicMock(side_effect=["old-press", "new-press"])
    release_hook = MagicMock(side_effect=["old-release", ValueError("Invalid key.")])
    unhook = MagicMock()
    monkeypatch.setattr(snap.keyboard, "on_press_key", press_hook, raising=False)
    monkeypatch.setattr(snap.keyboard, "on_release_key", release_hook, raising=False)
    monkeypatch.setattr(snap.keyboard, "unhook", unhook, raising=False)
    listener = snap.MultiPressHotkeyListener(_settings())

    result = listener.update_binding("not-a-key")

    assert result is False
    unhook.assert_called_once_with("new-press")
    assert listener.current_key == "shift"
    assert listener._press_hook == "old-press"
    assert listener._release_hook == "old-release"


def test_initial_hook_failure_disables_snap_without_crashing(monkeypatch):
    """A system hook failure must leave the application usable with snap disabled."""
    settings = _settings()
    monkeypatch.setattr(
        snap.keyboard,
        "on_press_key",
        MagicMock(side_effect=OSError("Hook unavailable.")),
        raising=False,
    )
    monkeypatch.setattr(snap.keyboard, "on_release_key", MagicMock(), raising=False)

    listener = snap.MultiPressHotkeyListener(settings)

    assert settings.enable_snap is False
    assert listener._press_hook is None
    assert listener._release_hook is None
