"""Regression tests for launch-at-login shortcut state reconciliation."""

from types import SimpleNamespace
from unittest.mock import MagicMock

from virelo.bridge.bridge import VireloBridge
from virelo.services.snap import SnapService


def _bridge_with_window(state):
    bridge = VireloBridge(state, SnapService(None))
    bridge.set_main_window(
        SimpleNamespace(
            action_run_at_startup=MagicMock(),
            action_minimize_on_exit=MagicMock(),
        )
    )
    return bridge


def test_failed_startup_enable_reconciles_the_persisted_toggle(monkeypatch):
    """A shortcut creation failure must not leave launch at login reported as enabled."""
    state = MagicMock()
    state.persist_immediate.return_value = {
        "ok": True,
        "applied": {"run_at_startup": False},
    }
    state.get_all.return_value = {"run_at_startup": False}
    bridge = _bridge_with_window(state)
    messages = []
    bridge.snap_status.connect(lambda message, timeout: messages.append((message, timeout)))
    monkeypatch.setattr(
        "virelo.app.window.create_startup_shortcut",
        MagicMock(side_effect=OSError("Shortcut creation failed.")),
    )
    monkeypatch.setattr("virelo.app.window.startup_shortcut_exists", lambda: False)
    applied = {"run_at_startup": True}

    bridge._apply_side_effects(applied)

    state.persist_immediate.assert_called_once_with({"run_at_startup": False})
    assert applied["run_at_startup"] is False
    bridge._main_window.action_run_at_startup.setChecked.assert_called_once_with(False)
    assert messages and "remains disabled" in messages[-1][0]


def test_failed_startup_disable_reconciles_to_an_existing_shortcut(monkeypatch):
    """A shortcut removal failure must keep launch at login reported as enabled."""
    state = MagicMock()
    state.persist_immediate.return_value = {
        "ok": True,
        "applied": {"run_at_startup": True},
    }
    state.get_all.return_value = {"run_at_startup": True}
    bridge = _bridge_with_window(state)
    monkeypatch.setattr(
        "virelo.app.window.remove_startup_shortcut",
        MagicMock(side_effect=PermissionError("Shortcut removal failed.")),
    )
    monkeypatch.setattr("virelo.app.window.startup_shortcut_exists", lambda: True)
    applied = {"run_at_startup": False}

    bridge._apply_side_effects(applied)

    state.persist_immediate.assert_called_once_with({"run_at_startup": True})
    assert applied["run_at_startup"] is True
    bridge._main_window.action_run_at_startup.setChecked.assert_called_once_with(True)
