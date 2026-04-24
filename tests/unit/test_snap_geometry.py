"""Tests for snap geometry and fullscreen detection (QUAL-03).

Tests calculate_snap_position and _rect_matches_monitor without requiring
Win32 APIs or a running desktop.
"""

from virelo.platform.win32_helpers import FULLSCREEN_TOLERANCE, _rect_matches_monitor
from virelo.services.snap import calculate_snap_position


# -- _rect_matches_monitor --


def test_rect_matches_monitor_exact():
    """Exact match should return True."""
    assert _rect_matches_monitor((0, 0, 1920, 1080), (0, 0, 1920, 1080)) is True


def test_rect_matches_monitor_within_tolerance():
    """Rect within FULLSCREEN_TOLERANCE pixels of monitor edges should match."""
    assert FULLSCREEN_TOLERANCE >= 2  # sanity check the constant
    assert _rect_matches_monitor((-2, -1, 1922, 1081), (0, 0, 1920, 1080)) is True


def test_rect_does_not_match_monitor():
    """Rect clearly not covering the monitor should return False."""
    assert _rect_matches_monitor((100, 100, 800, 600), (0, 0, 1920, 1080)) is False


def test_rect_matches_negative_coords():
    """Second monitor with negative coords should still match."""
    assert _rect_matches_monitor((-1920, 0, 0, 1080), (-1920, 0, 0, 1080)) is True


# -- calculate_snap_position --


def test_calculate_snap_position_76pct():
    """76% width/height on 1920x1080 starting at (0,0)."""
    x, y, w, h = calculate_snap_position(0, 0, 1920, 1080, 76, 76)
    assert w == 1920 * 76 // 100  # 1459
    assert h == 1080 * 76 // 100  # 820
    assert x == (1920 - w) // 2   # 230
    assert y == (1080 - h) // 2   # 130


def test_calculate_snap_position_50pct():
    """50% width/height on 1920x1080."""
    x, y, w, h = calculate_snap_position(0, 0, 1920, 1080, 50, 50)
    assert w == 960
    assert h == 540
    assert x == 480
    assert y == 270


def test_calculate_snap_position_100pct():
    """100% should fill the entire monitor."""
    x, y, w, h = calculate_snap_position(0, 0, 1920, 1080, 100, 100)
    assert (x, y, w, h) == (0, 0, 1920, 1080)


def test_calculate_snap_position_offset_monitor():
    """Monitor starting at x=1920 should shift the x position."""
    x, y, w, h = calculate_snap_position(1920, 0, 2560, 1440, 76, 76)
    assert w == 2560 * 76 // 100  # 1945
    assert h == 1440 * 76 // 100  # 1094
    # x includes monitor_left offset
    assert x == 1920 + (2560 - w) // 2
    assert y == (1440 - h) // 2
