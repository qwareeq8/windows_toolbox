"""Tests for SettingsState validation, draft model, and coercion (QUAL-03, STRUCT-05).

These tests run without PySide6/WebEngine by using the MockSettings fixture.
"""

from virelo.app.config import DEFAULTS


def test_apply_draft_validates_range(settings_state):
    """width_pct=150 exceeds max 100, should fail."""
    result = settings_state.apply_draft({"width_pct": 150})
    assert result["ok"] is False
    assert "must be between" in result["error"]


def test_apply_draft_validates_range_low(settings_state):
    """width_pct=5 is below min 10, should fail."""
    result = settings_state.apply_draft({"width_pct": 5})
    assert result["ok"] is False
    assert "must be between" in result["error"]


def test_apply_draft_rejects_unknown_keys(settings_state):
    """Unknown keys should be rejected."""
    result = settings_state.apply_draft({"nonexistent_key": "val"})
    assert result["ok"] is False
    assert "Unknown" in result["error"]


def test_apply_draft_coerces_types(settings_state):
    """String '5' should be coerced to int 5 for snap_presses."""
    result = settings_state.apply_draft({"snap_presses": "5"})
    assert result["ok"] is True
    assert result["applied"]["snap_presses"] == 5


def test_apply_draft_valid_values(settings_state):
    """Valid width and height should be accepted."""
    result = settings_state.apply_draft({"width_pct": 50, "height_pct": 80})
    assert result["ok"] is True
    assert result["applied"]["width_pct"] == 50
    assert result["applied"]["height_pct"] == 80


def test_get_all_returns_all_keys(settings_state):
    """get_all should return a dict containing every key from DEFAULTS."""
    result = settings_state.get_all()
    for key in DEFAULTS:
        assert key in result, f"Missing key: {key}"


def test_commit_draft_persists(settings_state):
    """After apply + commit, has_draft should be False."""
    settings_state.apply_draft({"width_pct": 50})
    result = settings_state.commit_draft()
    assert result["ok"] is True
    assert settings_state.has_draft is False


def test_discard_draft_clears(settings_state):
    """After apply + discard, has_draft should be False."""
    settings_state.apply_draft({"width_pct": 50})
    assert settings_state.has_draft is True
    settings_state.discard_draft()
    assert settings_state.has_draft is False


def test_has_draft_initially_false(settings_state):
    """Fresh SettingsState should have no draft."""
    assert settings_state.has_draft is False


def test_get_all_overlays_draft(settings_state):
    """After applying width_pct=50, get_all should reflect the draft value."""
    settings_state.apply_draft({"width_pct": 50})
    result = settings_state.get_all()
    assert result["width_pct"] == 50


def test_reset_to_defaults_clears_draft(settings_state):
    """reset_to_defaults should clear any pending draft."""
    settings_state.apply_draft({"width_pct": 50})
    settings_state.reset_to_defaults()
    assert settings_state.has_draft is False
