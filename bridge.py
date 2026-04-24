"""QWebChannel bridge between React frontend and Python backend.

Exposes a single VireloBridge QObject with narrow JSON-based Slot methods.
All inputs are validated. No broad Python object exposure.

The bridge is registered with QWebChannel as "bridge" so JavaScript accesses
it as channel.objects.bridge.

All slots return structured JSON payloads:
  Success: {"ok": true, "data": ...}
  Failure: {"ok": false, "error": "..."}
"""

import json
import logging
from typing import Optional

from PySide6 import QtCore
from PySide6.QtCore import QObject, Signal, Slot

from settings_state import SettingsState
from snap_service import SnapService

LOG = logging.getLogger("Virelo")


class VireloBridge(QObject):
    """Narrow JSON bridge between React UI and Python backend.

    All Slot methods accept/return JSON strings (or primitive types).
    Signals push updates from Python to React.

    Registration: QWebChannel.registerObject("bridge", self)
    JavaScript:   channel.objects.bridge.get_settings(callback)
    """

    # --- Signals (Python -> JS) ---
    settings_changed = Signal(str)    # JSON string of full settings dict
    theme_applied = Signal(str)       # "dark" or "light" (effective theme)
    snap_status = Signal(str, int)    # (message, timeout_ms)
    capture_status = Signal(str)      # "capturing", "done", "cancelled", "timeout"

    def __init__(self, settings_state: SettingsState, snap_service: SnapService,
                 parent: Optional[QObject] = None):
        super().__init__(parent)
        self._state = settings_state
        self._snap = snap_service

        # These are set by MainWindow after construction
        self._main_window = None
        self._capture_guard = None

    def set_main_window(self, mw):
        """Set MainWindow reference for theme/capture/startup operations."""
        self._main_window = mw

    def set_capture_guard(self, guard):
        """Set CaptureGuard for key capture gating."""
        self._capture_guard = guard

    # --- Settings Slots ---

    @Slot(result=str)
    def get_settings(self) -> str:
        """Return all settings as a structured JSON payload."""
        try:
            settings = self._state.get_all()
            return json.dumps({"ok": True, "data": settings})
        except Exception as e:
            LOG.exception("get_settings failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(str, result=str)
    def save_settings(self, json_str: str) -> str:
        """Store a partial settings update in draft (not persisted).

        Input: JSON dict of key-value pairs.
        Changes are staged in the draft model. Call commit_draft to persist.
        """
        try:
            data = json.loads(json_str)
            if not isinstance(data, dict):
                return json.dumps({"ok": False, "error": "Expected JSON object"})
            result = self._state.apply_draft(data)
            if result.get("ok"):
                # Push updated settings (with draft overlay) to frontend
                self.settings_changed.emit(self._state.get_json())
            return json.dumps(result)
        except json.JSONDecodeError as e:
            return json.dumps({"ok": False, "error": f"Invalid JSON: {e}"})
        except Exception as e:
            LOG.exception("save_settings failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(result=str)
    def commit_draft(self) -> str:
        """Persist draft settings to QSettings and apply side effects."""
        try:
            result = self._state.commit_draft()
            if result.get("ok"):
                self.settings_changed.emit(self._state.get_json())
                self._apply_side_effects(result.get("applied", {}))
            return json.dumps(result)
        except Exception as e:
            LOG.exception("commit_draft failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(result=str)
    def discard_draft(self) -> str:
        """Discard unsaved changes and push persisted settings to frontend."""
        try:
            self._state.discard_draft()
            self.settings_changed.emit(self._state.get_json())
            return json.dumps({"ok": True})
        except Exception as e:
            LOG.exception("discard_draft failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(result=str)
    def has_draft(self) -> str:
        """Return whether unsaved changes exist."""
        return json.dumps({"ok": True, "data": self._state.has_draft})

    @Slot(result=str)
    def reset_defaults(self) -> str:
        """Reset all settings to defaults. Returns new settings as structured payload."""
        try:
            new_settings = self._state.reset_to_defaults()
            self.settings_changed.emit(json.dumps({"ok": True, "data": new_settings}))
            # Apply all business logic side effects
            if self._main_window:
                self._main_window._update_snap_enabled_state()
                self._main_window._update_explorer_autosize_thread()
                if hasattr(self._main_window, "shift_mgr"):
                    self._main_window.shift_mgr.update_binding(new_settings["snap_key"])
                    self._main_window.shift_mgr.update_restore_key(new_settings["restore_key"])
                    self._main_window.shift_mgr.update_press_limit(new_settings["snap_presses"])
                self._main_window._apply_theme_mode(new_settings["theme"])
            return json.dumps({"ok": True, "data": new_settings})
        except Exception as e:
            LOG.exception("reset_defaults failed")
            return json.dumps({"ok": False, "error": str(e)})

    # --- Snap Slots ---

    @Slot(result=str)
    def test_snap(self) -> str:
        """Trigger a test snap on the active window."""
        try:
            result = self._snap.test_snap()
            msg = result.get("message", result.get("error", ""))
            timeout = 2000
            self.snap_status.emit(msg, timeout)
            return json.dumps(result)
        except Exception as e:
            LOG.exception("test_snap failed")
            return json.dumps({"ok": False, "error": str(e)})

    # --- Key Capture Slots ---

    @Slot(str, result=str)
    def capture_key(self, target: str) -> str:
        """Start key capture for 'snap' or 'restore' target."""
        if target not in ("snap", "restore"):
            return json.dumps({"ok": False, "error": f"Invalid target: {target}"})
        if self._main_window is None:
            return json.dumps({"ok": False, "error": "MainWindow not ready"})
        try:
            if target == "snap":
                self._main_window._start_key_capture()
            else:
                self._main_window._start_restore_key_capture()
            self.capture_status.emit("capturing")
            return json.dumps({"ok": True})
        except Exception as e:
            LOG.exception("capture_key failed")
            return json.dumps({"ok": False, "error": str(e)})

    # --- Theme Slots ---

    @Slot(str, result=str)
    def apply_theme(self, mode: str) -> str:
        """Apply theme mode: 'system', 'dark', or 'light'."""
        if mode not in ("system", "dark", "light"):
            return json.dumps({"ok": False, "error": f"Invalid theme mode: {mode}"})
        if self._main_window is None:
            return json.dumps({"ok": False, "error": "MainWindow not ready"})
        try:
            self._main_window._apply_theme_mode(mode)
            # Store in draft (not persisted until commit)
            self._state.apply_draft({"theme": mode})
            return json.dumps({"ok": True})
        except Exception as e:
            LOG.exception("apply_theme failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(result=str)
    def get_theme_mode(self) -> str:
        """Return current theme mode as structured JSON."""
        mode = "dark"
        if self._main_window:
            mode = getattr(self._main_window, "_theme_mode", "dark")
        return json.dumps({"ok": True, "data": mode})

    # --- Startup Slot ---

    @Slot(bool, result=str)
    def toggle_run_at_startup(self, enabled: bool) -> str:
        """Toggle run-at-startup shortcut."""
        if self._main_window is None:
            return json.dumps({"ok": False, "error": "MainWindow not ready"})
        try:
            self._main_window.action_run_at_startup.setChecked(enabled)
            self._main_window._toggle_run_at_startup()
            actual = self._main_window.action_run_at_startup.isChecked()
            return json.dumps({"ok": True, "data": {"enabled": actual}})
        except Exception as e:
            LOG.exception("toggle_run_at_startup failed")
            return json.dumps({"ok": False, "error": str(e)})

    @Slot(result=str)
    def get_launch_at_login(self) -> str:
        """Return current launch-at-login state as structured JSON."""
        if self._main_window:
            val = self._main_window.action_run_at_startup.isChecked()
            return json.dumps({"ok": True, "data": val})
        return json.dumps({"ok": True, "data": False})

    @Slot(result=str)
    def get_snap_enabled(self) -> str:
        """Return whether snap is currently enabled as structured JSON."""
        if self._main_window:
            val = bool(getattr(self._main_window, "snap_enabled", False))
            return json.dumps({"ok": True, "data": val})
        return json.dumps({"ok": True, "data": False})

    # --- Window Command Slot ---

    @Slot(str, result=str)
    def setWindowCommand(self, command: str) -> str:
        """Execute a window management command (minimize or close)."""
        if command not in ("minimize", "close"):
            return json.dumps({"ok": False, "error": f"Unknown command: {command}"})
        if self._main_window is None:
            return json.dumps({"ok": False, "error": "MainWindow not ready"})
        try:
            if command == "minimize":
                self._main_window.showMinimized()
            elif command == "close":
                self._main_window.close()
            return json.dumps({"ok": True})
        except Exception as e:
            LOG.exception("setWindowCommand(%s) failed", command)
            return json.dumps({"ok": False, "error": str(e)})

    # --- Internal helpers ---

    def _apply_side_effects(self, applied: dict):
        """Apply business logic side effects after settings are committed."""
        if not self._main_window:
            return
        mw = self._main_window

        if "enable_snap" in applied:
            mw.snap_enabled = bool(applied["enable_snap"])
            mw._update_snap_enabled_state()

        if "ex_auto_size" in applied:
            mw._update_explorer_autosize_thread()

        if "snap_presses" in applied and hasattr(mw, "shift_mgr"):
            mw.shift_mgr.update_press_limit(applied["snap_presses"])

        if "snap_key" in applied and hasattr(mw, "shift_mgr"):
            mw.shift_mgr.update_binding(applied["snap_key"])

        if "restore_key" in applied and hasattr(mw, "shift_mgr"):
            mw.shift_mgr.update_restore_key(applied["restore_key"])

        if "theme" in applied:
            mw._apply_theme_mode(applied["theme"])
