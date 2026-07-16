"""Shutdown safety checks for background folder-view registry tasks."""

import threading

from virelo.bridge.bridge import VireloBridge
from virelo.services.snap import SnapService


def test_wait_for_view_task_reports_timeout_and_completion(settings_state):
    """Shutdown callers can refuse quit until registry work truly finishes."""
    bridge = VireloBridge(settings_state, SnapService(None))
    release = threading.Event()
    thread = threading.Thread(target=release.wait)
    bridge._views_thread = thread
    thread.start()
    try:
        assert bridge.is_view_task_running() is True
        assert bridge.wait_for_view_task(timeout=0.01) is False
        assert bridge.is_view_task_running() is True
    finally:
        release.set()
        thread.join(1.0)

    assert bridge.wait_for_view_task(timeout=0.01) is True
    assert bridge.is_view_task_running() is False
