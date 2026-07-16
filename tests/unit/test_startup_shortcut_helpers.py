"""Tests for per-user Startup shortcut validation and rollback helpers."""

import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from virelo.app import window
from virelo.platform.startup import startup_shortcut_spec


def _isolated_shortcut_path(tmp_path: Path, monkeypatch) -> Path:
    monkeypatch.setenv("APPDATA", str(tmp_path))
    path = Path(window.get_startup_shortcut_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def test_shortcut_match_checks_the_current_target_arguments_and_working_directory(
    tmp_path, monkeypatch
):
    """A stale launch target must not count as an enabled startup shortcut."""
    path = _isolated_shortcut_path(tmp_path, monkeypatch)
    path.write_bytes(b"test shortcut")
    script = os.path.abspath(window.sys.argv[0])
    frozen = bool(getattr(window.sys, "frozen", False))
    target, arguments = startup_shortcut_spec(window.sys.executable, script, frozen)
    shortcut = SimpleNamespace(
        TargetPath=target,
        Arguments=arguments,
        WorkingDirectory=os.path.dirname(target if frozen else script),
    )
    shell = SimpleNamespace(CreateShortcut=MagicMock(return_value=shortcut))
    monkeypatch.setattr(window, "_ensure_dispatch", MagicMock(return_value=shell))

    assert window.startup_shortcut_matches_current_launch() is True

    shortcut.TargetPath = str(tmp_path / "stale" / "Virelo.exe")
    assert window.startup_shortcut_matches_current_launch() is False


def test_shortcut_snapshot_round_trips_exact_bytes(tmp_path, monkeypatch):
    """Rollback restores the prior shortcut exactly, including prior absence."""
    path = _isolated_shortcut_path(tmp_path, monkeypatch)
    original = b"original shortcut bytes"
    path.write_bytes(original)

    snapshot = window.read_startup_shortcut()
    path.write_bytes(b"replacement")
    window.restore_startup_shortcut(snapshot)

    assert path.read_bytes() == original

    window.restore_startup_shortcut(None)
    assert path.exists() is False


def test_failed_shortcut_save_does_not_delete_a_preexisting_link(tmp_path, monkeypatch):
    """The transaction owner, not the COM writer, controls rollback state."""
    path = _isolated_shortcut_path(tmp_path, monkeypatch)
    original = b"preexisting shortcut"
    path.write_bytes(original)
    shortcut = SimpleNamespace(Save=MagicMock(side_effect=OSError("COM save failed.")))
    shell = SimpleNamespace(CreateShortcut=MagicMock(return_value=shortcut))
    monkeypatch.setattr(window, "_ensure_dispatch", MagicMock(return_value=shell))

    with pytest.raises(OSError, match="COM save failed"):
        window.create_startup_shortcut()

    assert path.read_bytes() == original
