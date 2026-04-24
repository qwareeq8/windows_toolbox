"""Snap service wrapper for the VireloBridge.

Wraps the ShiftSnapRestore API so bridge.py can trigger snap actions
without directly depending on the ShiftSnapRestore class internals.
"""

import logging

LOG = logging.getLogger("Virelo")


def calculate_snap_position(
    monitor_left: int,
    monitor_top: int,
    monitor_width: int,
    monitor_height: int,
    width_pct: int,
    height_pct: int,
) -> tuple:
    """Calculate snap target position (x, y, w, h) for a resizable window."""
    w = monitor_width * width_pct // 100
    h = monitor_height * height_pct // 100
    x = monitor_left + (monitor_width - w) // 2
    y = monitor_top + (monitor_height - h) // 2
    return (x, y, w, h)


class SnapService:
    """Narrow API surface for snap/restore actions."""

    def __init__(self, shift_mgr):
        """Accept a ShiftSnapRestore instance (or None during early init)."""
        self._mgr = shift_mgr

    def set_manager(self, mgr):
        """Set or replace the ShiftSnapRestore instance."""
        self._mgr = mgr

    def test_snap(self) -> dict:
        """Trigger a test snap (same as pressing "Test snap" button)."""
        if self._mgr is None:
            return {"ok": False, "error": "Snap manager not initialized"}
        try:
            self._mgr.perform(False)
            return {"ok": True, "message": "Snap test applied to the active window."}
        except Exception as e:
            LOG.exception("test_snap failed")
            return {"ok": False, "error": str(e)}

    def update_binding(self, key: str):
        """Update the snap key binding."""
        if self._mgr:
            self._mgr.update_binding(key)

    def update_restore_key(self, key: str):
        """Update the restore key binding."""
        if self._mgr:
            self._mgr.update_restore_key(key)

    def update_press_limit(self, count: int):
        """Update the snap press count."""
        if self._mgr:
            self._mgr.update_press_limit(count)
