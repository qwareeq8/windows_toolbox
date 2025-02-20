# Windows Toolbox

This repository contains the complete Windows Toolbox application for snapping and restoring windows via a SHIFT triple-press mechanism, as well as an Explorer View Manager that adjusts Windows Explorer settings using hotkeys. The implementation is done in Python with PyQt5 and various Windows APIs. The code strictly follows the documentation style guidelines.

## Repository Structure

```
windows_toolbox/
├── windows_toolbox.py
└── README.md
```

## Project Overview

The Windows Toolbox application provides the following features:
- **SHIFT Triple-Press Snap & Restore:**  
  Detects three distinct SHIFT presses within a specified interval to either snap (center and resize) or restore the active window.
- **Explorer View Manager:**  
  Periodically polls Windows Explorer windows to adjust their view settings by sending hotkey commands.  
  Supports both repeated and one-shot approaches based on user settings.
- **User Settings:**  
  Settings such as hotkey, number of presses, interval, window size percentages, and Explorer view preferences are loaded from and saved to a JSON configuration file.

## How to Use

1. **Requirements:**  
   - Python 3  
   - PyQt5  
   - Additional modules: psutil, keyboard, ctypes, pywin32, win32com, etc.
2. **Execution:**  
   Run the application with:
   ```bash
   python windows_toolbox.py
   ```
3. **User Interaction:**  
   - Use SHIFT triple-press to snap/restore the active window.
   - The GUI allows you to adjust settings for both the snap & restore functionality and the Explorer View Manager.
