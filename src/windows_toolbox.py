#!/usr/bin/env python
#
# files: windows_toolbox.py
#
# revision history:
# 20251219 (XX): initial version
#
# This file contains a Python implementation of:
#   A Windows Toolbox application that implements SHIFT triple-press snapping/restoring of windows,
#   and an Explorer View Manager that adjusts Explorer settings (such as view mode, grouping,
#   and sorting) by sending hotkeys.
#
#------------------------------------------------------------------------------

# import required system modules
import os
import sys
import time
import json
import psutil
import threading
import keyboard
import ctypes
import win32gui
import win32api
import win32con
import win32com.client
from PyQt5 import QtCore, QtGui, QtWidgets

#------------------------------------------------------------------------------
#
# global variables are listed here
#
#------------------------------------------------------------------------------

# Get user32 from ctypes and define window-related constants
user32 = ctypes.windll.user32
SW_RESTORE = 9
MONITOR_DEFAULTTONEAREST = 2

# Define program-specific constants
PROGRAM_NAME = "windows_toolbox"   # Name of the program
SETTINGS_FILENAME = "snap_and_restore_settings.json"   # Settings file name
ICON_FILENAME = "icon.ico"         # Icon file name

# Define key mapping for hotkey processing
KEY_MAP = {
    "Left Shift": "shift",
    "Right Shift": "right shift",
    "Left Ctrl": "ctrl",
    "Right Ctrl": "right ctrl",
    "Left Alt": "alt",
    "Right Alt": "right alt",
    "A": "a",
    "B": "b",
    "C": "c"
}

#------------------------------------------------------------------------------
#
# Utility functions are defined here
#
#------------------------------------------------------------------------------

def resource_path(relative_path):
    """
    function: resource_path

    arguments:
      relative_path: (string) relative file path

    return:
      (string) absolute path to the resource

    description:
      Returns the absolute path to a resource, working whether the application is
      bundled (using _MEIPASS) or run as a script.
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(".")) if hasattr(sys, '_MEIPASS') else os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_settings_file_path():
    """
    function: get_settings_file_path

    arguments:
      none

    return:
      (string) the absolute path to the settings file

    description:
      Determines the location of the settings file depending on whether the program is bundled.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(os.path.dirname(sys.executable), SETTINGS_FILENAME)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(script_dir, SETTINGS_FILENAME)

def load_settings():
    """
    function: load_settings

    arguments:
      none

    return:
      (dict) settings loaded from file or default settings

    description:
      Loads settings from a JSON file if available; otherwise returns default settings.
    """
    default_settings = {
        "enable_snap_restore": True,
        "enable_explorer_view": True,
        "hotkey": "Left Shift",
        "presses": 3,
        "interval": 1050,
        "width_pct": 76,
        "height_pct": 76,
        "explorer_viewmode": 4,  # 4 => Details
        "explorer_sortcolumn": "System.ItemNameDisplay",
        "explorer_sortascending": True,
        "explorer_enablegrouping": False,
        "explorer_autosizecolumns": True,      # repeated approach
        "explorer_one_shot_ctrl_plus": True    # one-shot on folder changes
    }
    path = get_settings_file_path()
    if not os.path.exists(path):
        return default_settings
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for k in default_settings:
            if k not in data:
                data[k] = default_settings[k]
        return data
    except:
        return default_settings

def save_settings(settings):
    """
    function: save_settings

    arguments:
      settings: (dict) settings to save

    return:
      none

    description:
      Saves the provided settings dictionary to a JSON file.
    """
    path = get_settings_file_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)

def is_explorer_foreground():
    """
    function: is_explorer_foreground

    arguments:
      none

    return:
      (bool) True if the current foreground window belongs to explorer.exe, False otherwise

    description:
      Checks if the current foreground window is part of Windows Explorer.
    """
    fg_hwnd = user32.GetForegroundWindow()
    if not fg_hwnd:
        return False
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(fg_hwnd, ctypes.byref(pid))
    try:
        proc = psutil.Process(pid.value)
        return (proc.name().lower() == "explorer.exe")
    except:
        return False

def send_ctrl_plus():
    """
    function: send_ctrl_plus

    arguments:
      none

    return:
      none

    description:
      Simulates pressing Ctrl + numeric keypad '+' once.
    """
    keyboard.press_and_release('ctrl+add')

def unmaximize_if_needed(hwnd):
    """
    function: unmaximize_if_needed

    arguments:
      hwnd: (handle) window handle

    return:
      none

    description:
      If the window is maximized, restores it to normal size.
    """
    if user32.IsZoomed(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)

def get_window_hash(hwnd):
    """
    function: get_window_hash

    arguments:
      hwnd: (handle) window handle

    return:
      (string) a unique hash representing the window

    description:
      Constructs a unique identifier for a window based on its process ID,
      class name, and window text.
    """
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    pid_val = pid.value
    class_name = win32gui.GetClassName(hwnd)
    window_text = win32gui.GetWindowText(hwnd)
    return f"{pid_val}_{class_name}_{window_text}"

def find_child_window(parent_hwnd, child_class):
    """
    function: find_child_window

    arguments:
      parent_hwnd: (handle) parent window handle
      child_class: (string) class name of child window to find

    return:
      (handle) child window handle if found, else None

    description:
      Searches for a child window with a specific class name under the given parent.
    """
    direct = win32gui.FindWindowEx(parent_hwnd, 0, child_class, None)
    if direct:
        return direct

    found = []
    def enum_callback(hwnd, _):
        cname = win32gui.GetClassName(hwnd)
        if cname.lower() == child_class.lower():
            found.append(hwnd)
    win32gui.EnumChildWindows(parent_hwnd, enum_callback, None)
    return found[0] if found else None

#------------------------------------------------------------------------------
#
# SHIFT TRIPLE-PRESS logic is implemented in the following class
#
#------------------------------------------------------------------------------

class ShiftTriplePress(QtCore.QObject):
    """
    Class: ShiftTriplePress

    arguments:
     settings: (dict) configuration settings for snap & restore functionality
     parent: (QObject) parent object (default None)

    description:
     Implements SHIFT triple-press logic that avoids counting a continuous hold as multiple presses.
     Once the required number of distinct SHIFT presses within a specified interval is detected,
     the active window is either snapped (centered and resized) or restored to its original size.
    """
    def __init__(self, settings, parent=None):
        """
        method: __init__

        arguments:
         settings: (dict) configuration settings
         parent: (QObject) parent object (default None)

        return:
         none

        description:
         Initializes the ShiftTriplePress instance, hooks the SHIFT key, and sets initial parameters.
        """
        super().__init__(parent)
        self.settings = settings
        self.shift_held = False
        self.press_times = []
        self.required_presses = self.settings["presses"]
        self.interval_s = self.settings["interval"] / 1000.0

        # Store original window sizes for restoration
        self.original_sizes = {}

        # Hook SHIFT key events
        keyboard.hook_key('shift', self.on_shift_event, suppress=False)

    def on_shift_event(self, event):
        """
        method: on_shift_event

        arguments:
         event: (keyboard event) the key event

        return:
         none

        description:
         Processes SHIFT key events to detect distinct presses and triggers snap or restore.
        """
        if not self.settings.get("enable_snap_restore", False):
            return

        if event.event_type == 'down':
            if not self.shift_held:
                self.shift_held = True
                now = time.time()
                # Remove old press times outside the interval
                self.press_times = [t for t in self.press_times if now - t <= self.interval_s]
                self.press_times.append(now)
                if len(self.press_times) == self.required_presses:
                    self.press_times.clear()
                    if keyboard.is_pressed('ctrl'):
                        self.center_original_size()
                    else:
                        self.center_and_resize_window()
        elif event.event_type == 'up':
            self.shift_held = False

    def center_and_resize_window(self):
        """
        method: center_and_resize_window

        arguments:
         none

        return:
         none

        description:
         Centers and resizes the currently active window based on configured width and height percentages.
        """
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return
        unmaximize_if_needed(hwnd)

        rect = ctypes.wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        o_left, o_top = rect.left, rect.top
        o_width = rect.right - rect.left
        o_height = rect.bottom - rect.top

        w_hash = get_window_hash(hwnd)
        if w_hash not in self.original_sizes:
            self.original_sizes[w_hash] = (o_left, o_top, o_width, o_height)

        width_pct = self.settings["width_pct"] / 100.0
        height_pct = self.settings["height_pct"] / 100.0

        monitor = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
        if not monitor:
            return
        mon_info = win32api.GetMonitorInfo(monitor)
        (m_left, m_top, m_right, m_bottom) = mon_info["Monitor"]
        m_w = m_right - m_left
        m_h = m_bottom - m_top

        new_w = int(m_w * width_pct)
        new_h = int(m_h * height_pct)
        new_left = m_left + (m_w - new_w) // 2
        new_top = m_top + (m_h - new_h) // 2

        user32.MoveWindow(hwnd, new_left, new_top, new_w, new_h, True)

    def center_original_size(self):
        """
        method: center_original_size

        arguments:
         none

        return:
         none

        description:
         Restores the active window to its original size and centers it.
        """
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return
        unmaximize_if_needed(hwnd)

        w_hash = get_window_hash(hwnd)
        if w_hash not in self.original_sizes:
            return
        (o_left, o_top, o_w, o_h) = self.original_sizes[w_hash]

        monitor = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
        if not monitor:
            return
        mon_info = win32api.GetMonitorInfo(monitor)
        (m_left, m_top, m_right, m_bottom) = mon_info["Monitor"]
        m_w = m_right - m_left
        m_h = m_bottom - m_top

        new_left = m_left + (m_w - o_w) // 2
        new_top = m_top + (m_h - o_h) // 2
        user32.MoveWindow(hwnd, new_left, new_top, o_w, o_h, True)

#------------------------------------------------------------------------------
#
# Explorer View Manager is implemented in the following class
#
#------------------------------------------------------------------------------

class ExplorerViewManager(QtCore.QObject):
    """
    Class: ExplorerViewManager

    arguments:
     settings: (dict) configuration settings for Explorer view adjustments
     parent: (QObject) parent object (default None)

    description:
     Implements a polling approach that, every 4 seconds, either sends repeated 'Ctrl+ +'
     commands if Explorer is in Details view or performs a one-shot 'Ctrl+ +' if a folder
     change is detected.
    """
    def __init__(self, settings, parent=None):
        """
        method: __init__

        arguments:
         settings: (dict) configuration settings
         parent: (QObject) parent object (default None)

        return:
         none

        description:
         Initializes the ExplorerViewManager instance and starts the polling timer.
        """
        super().__init__(parent)
        self.settings = settings
        self.enabled = self.settings["enable_explorer_view"]
        self.last_paths = {}  # Dictionary: {hwnd: last_known_path}

        self.timer = QtCore.QTimer()
        self.timer.setInterval(4000)  # Poll every 4 seconds
        self.timer.timeout.connect(self.poll_explorer)
        if self.enabled:
            self.timer.start()

    def set_settings(self, settings):
        """
        method: set_settings

        arguments:
         settings: (dict) updated configuration settings

        return:
         none

        description:
         Updates the settings and reconfigures the manager's enabled state.
        """
        self.settings = settings
        self.set_enabled(self.settings["enable_explorer_view"])

    def set_enabled(self, enable):
        """
        method: set_enabled

        arguments:
         enable: (bool) whether to enable the Explorer view manager

        return:
         none

        description:
         Enables or disables the polling timer and clears stored paths if disabled.
        """
        self.enabled = enable
        if enable:
            if not self.timer.isActive():
                self.timer.start()
        else:
            self.timer.stop()
            self.last_paths.clear()

    def poll_explorer(self):
        """
        method: poll_explorer

        arguments:
         none

        return:
         none

        description:
         Polls Explorer windows and sends hotkey commands based on the configured approach.
        """
        if not self.enabled:
            return

        shell = win32com.client.Dispatch("Shell.Application")
        for window in shell.Windows():
            if not window:
                continue
            try:
                name_lower = window.Name.lower()
                if "explorer" not in name_lower:
                    continue
            except:
                continue

            doc = window.Document
            if not doc:
                continue

            hwnd = window.HWND
            # One-shot approach for folder changes
            if self.settings.get("explorer_one_shot_ctrl_plus", False):
                new_path = ""
                try:
                    new_path = doc.Folder.Self.Path
                except:
                    pass
                old_path = self.last_paths.get(hwnd, None)
                if new_path and new_path != old_path:
                    self.last_paths[hwnd] = new_path
                    if is_explorer_foreground():
                        send_ctrl_plus()

            # Repeated approach if enabled
            if self.settings.get("explorer_autosizecolumns", False):
                try:
                    doc.CurrentViewMode = int(self.settings.get("explorer_viewmode", 4))
                except:
                    pass

                if not self.settings.get("explorer_enablegrouping", False):
                    try:
                        doc.GroupBy = "System.Null"
                    except:
                        pass
                else:
                    try:
                        doc.GroupBy = self.settings["explorer_sortcolumn"]
                    except:
                        pass

                try:
                    doc.SortColumns = self.settings["explorer_sortcolumn"]
                    doc.SortAscending = bool(self.settings["explorer_sortascending"])
                except:
                    pass

                mode = 0
                try:
                    mode = doc.CurrentViewMode
                except:
                    pass
                if mode == 4:
                    for _ in range(5):
                        send_ctrl_plus()
                        time.sleep(0.1)

                try:
                    doc.Refresh()
                except:
                    pass

#------------------------------------------------------------------------------
#
# Main UI is implemented in the following class
#
#------------------------------------------------------------------------------

class ToolboxMainWindow(QtWidgets.QMainWindow):
    """
    Class: ToolboxMainWindow

    arguments:
     none

    description:
     Main window that combines the SHIFT triple-press snap & restore logic with the Explorer View Manager.
    """
    def __init__(self):
        """
        method: __init__

        arguments:
         none

        return:
         none

        description:
         Initializes the main window, loads settings, sets up UI components, and configures system tray functionality.
        """
        super().__init__()
        icon_file = resource_path(ICON_FILENAME)
        self.setWindowIcon(QtGui.QIcon(icon_file))
        self.setWindowTitle("Windows Toolbox (Repeated Ctrl+ +)")
        self.setGeometry(100, 100, 700, 640)
        self.settings = load_settings()

        # Initialize SHIFT triple-press manager
        self.shift_manager = ShiftTriplePress(self.settings)
        # Initialize Explorer view manager
        self.explorer_manager = ExplorerViewManager(self.settings)

        main_layout = QtWidgets.QVBoxLayout()

        # Snap & Restore checkbox
        self.snap_checkbox = QtWidgets.QCheckBox("Enable Snap & Restore (SHIFT triple-press)")
        self.snap_checkbox.setChecked(self.settings["enable_snap_restore"])
        self.snap_checkbox.stateChanged.connect(self.toggle_snap)
        main_layout.addWidget(self.snap_checkbox)

        snap_group = QtWidgets.QGroupBox("Snap & Restore Settings")
        snap_layout = QtWidgets.QVBoxLayout()

        # Presses slider and label
        press_layout = QtWidgets.QHBoxLayout()
        self.press_label = QtWidgets.QLabel(f"Times Pressed: {self.settings['presses']}")
        self.press_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.press_slider.setRange(1, 10)
        self.press_slider.setValue(self.settings["presses"])
        self.press_slider.valueChanged.connect(self.update_press_label)
        press_layout.addWidget(self.press_label)
        press_layout.addWidget(self.press_slider)
        snap_layout.addLayout(press_layout)

        # Interval slider and label
        interval_layout = QtWidgets.QHBoxLayout()
        self.interval_label = QtWidgets.QLabel(f"Press Interval (ms): {self.settings['interval']}")
        self.interval_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.interval_slider.setRange(100, 3000)
        self.interval_slider.setValue(self.settings["interval"])
        self.interval_slider.valueChanged.connect(self.update_interval_label)
        interval_layout.addWidget(self.interval_label)
        interval_layout.addWidget(self.interval_slider)
        snap_layout.addLayout(interval_layout)

        # Width percentage slider and label
        horiz_layout = QtWidgets.QHBoxLayout()
        self.horiz_label = QtWidgets.QLabel(f"Width Percentage: {self.settings['width_pct']}%")
        self.horiz_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.horiz_slider.setRange(10, 100)
        self.horiz_slider.setValue(self.settings["width_pct"])
        self.horiz_slider.valueChanged.connect(self.update_horiz_label)
        horiz_layout.addWidget(self.horiz_label)
        horiz_layout.addWidget(self.horiz_slider)
        snap_layout.addLayout(horiz_layout)

        # Height percentage slider and label
        vert_layout = QtWidgets.QHBoxLayout()
        self.vert_label = QtWidgets.QLabel(f"Height Percentage: {self.settings['height_pct']}%")
        self.vert_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.vert_slider.setRange(10, 100)
        self.vert_slider.setValue(self.settings["height_pct"])
        self.vert_slider.valueChanged.connect(self.update_vert_label)
        vert_layout.addWidget(self.vert_label)
        vert_layout.addWidget(self.vert_slider)
        snap_layout.addLayout(vert_layout)

        snap_group.setLayout(snap_layout)
        main_layout.addWidget(snap_group)

        # Explorer Manager checkbox
        self.explorer_checkbox = QtWidgets.QCheckBox("Enable Explorer View Manager")
        self.explorer_checkbox.setChecked(self.settings["enable_explorer_view"])
        self.explorer_checkbox.stateChanged.connect(self.toggle_explorer)
        main_layout.addWidget(self.explorer_checkbox)

        explorer_group = QtWidgets.QGroupBox("Explorer View Settings")
        explorer_layout = QtWidgets.QVBoxLayout()

        # View mode selection
        viewmode_layout = QtWidgets.QHBoxLayout()
        viewmode_label = QtWidgets.QLabel("View Mode:")
        self.viewmode_combo = QtWidgets.QComboBox()
        possible_view_modes = [
            ("Large Icons (1)", 1),
            ("Small Icons (2)", 2),
            ("List (3)", 3),
            ("Details (4)", 4),
            ("Tiles (5)", 5),
            ("Content (7)", 7),
        ]
        for label, val in possible_view_modes:
            self.viewmode_combo.addItem(label, val)
        current_vm = self.settings["explorer_viewmode"]
        idx = 0
        for i in range(self.viewmode_combo.count()):
            if self.viewmode_combo.itemData(i) == current_vm:
                idx = i
                break
        self.viewmode_combo.setCurrentIndex(idx)
        viewmode_layout.addWidget(viewmode_label)
        viewmode_layout.addWidget(self.viewmode_combo)
        explorer_layout.addLayout(viewmode_layout)

        # Sort column configuration
        sortcol_layout = QtWidgets.QHBoxLayout()
        sortcol_label = QtWidgets.QLabel("Sort By Column (PropertyKey):")
        self.sortcol_edit = QtWidgets.QLineEdit(self.settings["explorer_sortcolumn"])
        sortcol_layout.addWidget(sortcol_label)
        sortcol_layout.addWidget(self.sortcol_edit)
        explorer_layout.addLayout(sortcol_layout)

        # Sort ascending option
        self.sortasc_checkbox = QtWidgets.QCheckBox("Sort Ascending (unchecked => descending)")
        self.sortasc_checkbox.setChecked(self.settings["explorer_sortascending"])
        explorer_layout.addWidget(self.sortasc_checkbox)

        # Grouping option
        self.grouping_checkbox = QtWidgets.QCheckBox("Enable Grouping")
        self.grouping_checkbox.setChecked(self.settings["explorer_enablegrouping"])
        explorer_layout.addWidget(self.grouping_checkbox)

        # Repeated approach option
        self.autosize_checkbox = QtWidgets.QCheckBox("Repeated Ctrl+ + calls if in Details view")
        self.autosize_checkbox.setChecked(self.settings["explorer_autosizecolumns"])
        explorer_layout.addWidget(self.autosize_checkbox)

        # One-shot approach option
        self.oneshot_checkbox = QtWidgets.QCheckBox("One-Shot Ctrl+ + on folder changes")
        self.oneshot_checkbox.setChecked(self.settings["explorer_one_shot_ctrl_plus"])
        explorer_layout.addWidget(self.oneshot_checkbox)

        explorer_group.setLayout(explorer_layout)
        main_layout.addWidget(explorer_group)

        save_btn = QtWidgets.QPushButton("Save All Settings")
        save_btn.clicked.connect(self.save_all)
        main_layout.addWidget(save_btn)

        info_lbl = QtWidgets.QLabel(
            "Shift Triple-Press:\n"
            "  - Holding SHIFT doesn't cause multiple presses.\n"
            "  - 3 distinct presses in the interval => snap or restore.\n\n"
            "Explorer Manager:\n"
            "  - If 'Repeated Ctrl+ + calls' is ON, we spam Ctrl+ + multiple times\n"
            "    if Explorer is forced to Details view.\n"
            "  - If 'One-Shot' is ON, we do a single Ctrl+ + when a folder change is detected\n"
            "    (only if Explorer is in foreground).\n"
            "Some folders or templates may ignore these requests.\n"
        )
        main_layout.addWidget(info_lbl)

        container = QtWidgets.QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Setup system tray functionality
        icon_ = QtGui.QIcon(icon_file)
        self.tray_icon = QtWidgets.QSystemTrayIcon(self)
        self.tray_icon.setIcon(icon_)
        self.tray_icon.setToolTip(PROGRAM_NAME)
        tray_menu = QtWidgets.QMenu()
        open_action = tray_menu.addAction("Open " + PROGRAM_NAME)
        open_action.triggered.connect(self.showNormal)
        quit_action = tray_menu.addAction("Quit")
        quit_action.triggered.connect(self.close_app)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_double_click)
        self.tray_icon.show()

    def toggle_snap(self, state):
        self.settings["enable_snap_restore"] = (state == QtCore.Qt.Checked)

    def toggle_explorer(self, state):
        self.settings["enable_explorer_view"] = (state == QtCore.Qt.Checked)
        self.explorer_manager.set_settings(self.settings)

    def update_press_label(self):
        val = self.press_slider.value()
        self.press_label.setText(f"Times Pressed: {val}")

    def update_interval_label(self):
        val = self.interval_slider.value()
        self.interval_label.setText(f"Press Interval (ms): {val}")

    def update_horiz_label(self):
        val = self.horiz_slider.value()
        self.horiz_label.setText(f"Width Percentage: {val}%")

    def update_vert_label(self):
        val = self.vert_slider.value()
        self.vert_label.setText(f"Height Percentage: {val}%")

    def save_all(self):
        self.settings["enable_snap_restore"] = self.snap_checkbox.isChecked()
        self.settings["presses"] = self.press_slider.value()
        self.settings["interval"] = self.interval_slider.value()
        self.settings["width_pct"] = self.horiz_slider.value()
        self.settings["height_pct"] = self.vert_slider.value()
        self.settings["enable_explorer_view"] = self.explorer_checkbox.isChecked()
        self.settings["explorer_viewmode"] = self.viewmode_combo.currentData()
        self.settings["explorer_sortcolumn"] = self.sortcol_edit.text().strip()
        self.settings["explorer_sortascending"] = self.sortasc_checkbox.isChecked()
        self.settings["explorer_enablegrouping"] = self.grouping_checkbox.isChecked()
        self.settings["explorer_autosizecolumns"] = self.autosize_checkbox.isChecked()
        self.settings["explorer_one_shot_ctrl_plus"] = self.oneshot_checkbox.isChecked()

        save_settings(self.settings)

        # Update SHIFT manager parameters
        self.shift_manager.required_presses = self.settings["presses"]
        self.shift_manager.interval_s = self.settings["interval"] / 1000.0

        # Update Explorer manager settings
        self.explorer_manager.set_settings(self.settings)

    def on_tray_icon_double_click(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.DoubleClick:
            self.showNormal()
            self.activateWindow()

    def close_app(self):
        keyboard.unhook_all()
        self.tray_icon.hide()
        QtWidgets.QApplication.quit()

    def closeEvent(self, event):
        event.ignore()
        self.showMinimized()

    def changeEvent(self, event):
        if event.type() == QtCore.QEvent.WindowStateChange:
            if self.isMinimized():
                QtCore.QTimer.singleShot(0, self.hide)
        super().changeEvent(event)

class ToolboxApp(QtWidgets.QApplication):
    """
    Class: ToolboxApp

    arguments:
     argv: (list) command line arguments

    description:
     Main application class that creates and shows the ToolboxMainWindow.
    """
    def __init__(self, argv):
        """
        method: __init__

        arguments:
         argv: (list) command line arguments

        return:
         none

        description:
         Initializes the ToolboxApp and creates the main window.
        """
        super().__init__(argv)
        self.main_window = ToolboxMainWindow()
        self.main_window.show()

def main():
    """
    method: main

    arguments:
     none

    return:
     none

    description:
     Main routine to initialize and execute the ToolboxApp.
    """
    app = ToolboxApp(sys.argv)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
#
# end of file
