"""Regression tests for transactional settings side effects."""

import json
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from virelo.bridge.bridge import VireloBridge
from virelo.services.snap import SnapService


class _Listener:
    """Track hotkey changes while providing deterministic failure injection."""

    def __init__(self, key: str = "shift") -> None:
        self.current_key = key
        self.restore_key = "ctrl"
        self.binding_calls: list[str] = []
        self.fail_next_binding = False

    def update_binding(self, key: str) -> bool:
        self.binding_calls.append(key)
        if self.fail_next_binding:
            self.fail_next_binding = False
            return False
        self.current_key = key
        return True

    def update_restore_key(self, key: str) -> None:
        self.restore_key = key

    def update_press_limit(self, count: int) -> None:
        del count


def _bridge_with_window(settings_state, listener: _Listener | None = None):
    bridge = VireloBridge(settings_state, SnapService(None))
    listener = listener or _Listener()
    window = SimpleNamespace(
        _hotkey_listener=listener,
        action_run_at_startup=MagicMock(),
        action_minimize_on_exit=MagicMock(),
        _update_snap_enabled_state=MagicMock(),
        _update_explorer_autosize_thread=MagicMock(),
        _apply_theme_mode=MagicMock(),
        snap_enabled=True,
        minimize_to_tray_on_exit=True,
    )
    bridge.set_main_window(window)
    return bridge, window, listener


def _mock_shortcut_state(monkeypatch, contents: bytes | None, matches: bool = False):
    """Isolate shortcut transactions from the real per-user Startup folder."""
    restore = MagicMock()
    monkeypatch.setattr("virelo.app.window.read_startup_shortcut", lambda: contents)
    monkeypatch.setattr(
        "virelo.app.window.startup_shortcut_matches_current_launch",
        lambda: matches,
    )
    monkeypatch.setattr("virelo.app.window.restore_startup_shortcut", restore)
    return restore


def test_hotkey_failure_keeps_the_setting_in_draft(settings_state, mock_settings):
    """A rejected hook must prevent persistence and leave the edit retryable."""
    listener = _Listener()
    listener.fail_next_binding = True
    bridge, _, _ = _bridge_with_window(settings_state, listener)
    settings_state.apply_draft({"snap_key": "alt"})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is False
    assert mock_settings.snap_key == "shift"
    assert settings_state.pending_changes == {"snap_key": "alt"}
    assert listener.current_key == "shift"
    assert listener.binding_calls == ["alt", "shift"]


def test_startup_creation_failure_prevents_persistence(settings_state, mock_settings, monkeypatch):
    """A failed shortcut creation must not persist an enabled toggle."""
    bridge, _, _ = _bridge_with_window(settings_state)
    create = MagicMock(side_effect=OSError("Shortcut creation failed."))
    restore = _mock_shortcut_state(monkeypatch, None)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", create)
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", MagicMock())
    settings_state.apply_draft({"run_at_startup": True})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is False
    assert mock_settings.run_at_startup is False
    assert settings_state.pending_changes == {"run_at_startup": True}
    create.assert_called_once_with()
    restore.assert_called_once_with(None)


def test_settings_save_failure_rolls_back_a_created_shortcut(
    settings_state, mock_settings, monkeypatch
):
    """A persistence failure must undo a shortcut created during preparation."""
    bridge, _, _ = _bridge_with_window(settings_state)
    create = MagicMock()
    restore = _mock_shortcut_state(monkeypatch, None)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", create)
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", MagicMock())
    calls = 0

    def fail_once_then_allow_rollback() -> None:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise OSError("Settings save failed.")

    mock_settings.save = fail_once_then_allow_rollback
    settings_state.apply_draft({"run_at_startup": True})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is False
    assert mock_settings.run_at_startup is False
    assert settings_state.pending_changes == {"run_at_startup": True}
    create.assert_called_once_with()
    restore.assert_called_once_with(None)
    assert calls == 2


def test_successful_startup_change_applies_the_external_effect_once(
    settings_state, mock_settings, monkeypatch
):
    """A successful commit must not create the startup shortcut twice."""
    bridge, window, _ = _bridge_with_window(settings_state)
    create = MagicMock()
    restore = _mock_shortcut_state(monkeypatch, None)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", create)
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", MagicMock())
    settings_state.apply_draft({"run_at_startup": True})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is True
    assert mock_settings.run_at_startup is True
    assert settings_state.has_draft is False
    create.assert_called_once_with()
    restore.assert_not_called()
    window.action_run_at_startup.setChecked.assert_called_once_with(True)


def test_failed_reset_restores_the_preexisting_draft(settings_state, mock_settings, monkeypatch):
    """A failed reset must preserve persisted values and the user's prior draft."""
    bridge, _, _ = _bridge_with_window(settings_state)
    _mock_shortcut_state(monkeypatch, None)
    mock_settings.width_pct = 70
    settings_state.apply_draft({"width_pct": 50})
    calls = 0

    def fail_once_then_allow_rollback() -> None:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise OSError("Settings save failed.")

    mock_settings.save = fail_once_then_allow_rollback

    result = json.loads(bridge.reset_defaults())

    assert result["ok"] is False
    assert mock_settings.width_pct == 70
    assert settings_state.pending_changes == {"width_pct": 50}
    assert settings_state.get_all()["width_pct"] == 50
    assert calls == 2


def test_immediate_startup_failure_does_not_persist(settings_state, mock_settings, monkeypatch):
    """The tray toggle must fail before changing persisted settings."""
    bridge, _, _ = _bridge_with_window(settings_state)
    create = MagicMock(side_effect=PermissionError("Startup folder is unavailable."))
    restore = _mock_shortcut_state(monkeypatch, None)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", create)
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", MagicMock())
    mock_settings.save = MagicMock()

    with pytest.raises(PermissionError, match="Startup folder is unavailable"):
        bridge.persist_immediate_settings({"run_at_startup": True})

    assert mock_settings.run_at_startup is False
    mock_settings.save.assert_not_called()
    restore.assert_called_once_with(None)


def test_stale_startup_shortcut_is_refreshed(settings_state, mock_settings, monkeypatch):
    """An existing shortcut with a stale target must be replaced on enable."""
    bridge, _, _ = _bridge_with_window(settings_state)
    create = MagicMock()
    restore = _mock_shortcut_state(monkeypatch, b"stale shortcut", matches=False)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", create)
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", MagicMock())
    settings_state.apply_draft({"run_at_startup": True})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is True
    assert mock_settings.run_at_startup is True
    create.assert_called_once_with()
    restore.assert_not_called()


def test_broken_startup_shortcut_can_still_be_disabled(settings_state, mock_settings, monkeypatch):
    """Disabling must remove an unreadable link without trying to inspect it."""
    mock_settings.run_at_startup = True
    bridge, _, _ = _bridge_with_window(settings_state)
    remove = MagicMock()
    match = MagicMock(side_effect=OSError("The link is corrupt."))
    restore = MagicMock()
    monkeypatch.setattr("virelo.app.window.read_startup_shortcut", lambda: b"corrupt link")
    monkeypatch.setattr("virelo.app.window.startup_shortcut_matches_current_launch", match)
    monkeypatch.setattr("virelo.app.window.restore_startup_shortcut", restore)
    monkeypatch.setattr("virelo.app.window.create_startup_shortcut", MagicMock())
    monkeypatch.setattr("virelo.app.window.remove_startup_shortcut", remove)
    settings_state.apply_draft({"run_at_startup": False})

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is True
    assert mock_settings.run_at_startup is False
    match.assert_not_called()
    remove.assert_called_once_with()
    restore.assert_not_called()


def test_later_side_effects_continue_after_one_component_fails(settings_state):
    """One post-save failure must not prevent unrelated UI updates."""
    bridge, window, _ = _bridge_with_window(settings_state)
    window._update_explorer_autosize_thread.side_effect = OSError("Worker unavailable.")
    messages: list[tuple[str, int]] = []
    bridge.snap_status.connect(lambda message, timeout: messages.append((message, timeout)))
    settings_state.apply_draft(
        {
            "ex_auto_size": True,
            "theme": "dark",
            "minimize_to_tray": False,
        }
    )

    result = json.loads(bridge.commit_draft())

    assert result["ok"] is True
    window._apply_theme_mode.assert_called_once_with("dark")
    window.action_minimize_on_exit.setChecked.assert_called_once_with(False)
    assert messages
    assert "Explorer column auto-size" in messages[-1][0]


def test_reset_returns_component_warnings_without_hiding_them(settings_state, monkeypatch):
    """Reset must not replace a post-save warning with a success message."""
    bridge, window, _ = _bridge_with_window(settings_state)
    _mock_shortcut_state(monkeypatch, None)
    window._update_explorer_autosize_thread.side_effect = OSError("Worker unavailable.")
    messages: list[tuple[str, int]] = []
    bridge.snap_status.connect(lambda message, timeout: messages.append((message, timeout)))

    result = json.loads(bridge.reset_defaults())

    assert result["ok"] is True
    assert result["warnings"] == ["Explorer column auto-size"]
    assert messages
    assert "Explorer column auto-size" in messages[-1][0]
    assert all(message != "Defaults loaded." for message, _ in messages)
