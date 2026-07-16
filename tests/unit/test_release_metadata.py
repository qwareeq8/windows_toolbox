"""Tests that keep Windows installer metadata aligned with the application."""

import re
from pathlib import Path

from virelo.app.__main__ import MUTEX_NAMES

PROJECT_ROOT = Path(__file__).resolve().parents[2]
INSTALLER_SCRIPT = PROJECT_ROOT / "installer" / "virelo.iss"


def _read_setup_value(name: str) -> str:
    """Return one scalar value from the Inno Setup script."""
    content = INSTALLER_SCRIPT.read_text(encoding="utf-8")
    match = re.search(rf"^{re.escape(name)}=(.+)$", content, re.MULTILINE)
    assert match is not None, f"{name} is missing from the installer metadata."
    return match.group(1).strip()


def test_installer_mutex_matches_application_mutex():
    """The installer must detect current and shipped legacy application mutexes."""
    assert tuple(_read_setup_value("AppMutex").split(",")) == MUTEX_NAMES


def test_installer_requires_the_qt_supported_windows_floor():
    """The installer must reject Windows releases unsupported by Qt 6.11."""
    assert _read_setup_value("MinVersion") == "10.0.17763"


def test_admin_installer_does_not_mutate_per_user_areas():
    """A machine-wide installer must leave each account's files to that account."""
    content = INSTALLER_SCRIPT.read_text(encoding="utf-8").lower()
    assert "{localappdata}" not in content
    assert "{userstartup}" not in content
    assert "{userappdata}" not in content
