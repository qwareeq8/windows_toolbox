# Troubleshooting Virelo

## Build Issues

### PyInstaller fails with PySide6 import error in Virelo.spec

**Cause:** `virelo/app/config.py` was imported directly in the spec file. The spec runs in
PyInstaller's analysis context where PySide6 may not be importable.

**Fix:** The spec file uses regex to parse `APP_VERSION` from `config.py`. Never add
`from virelo.app.config import ...` to `Virelo.spec`.

### PowerShell build script reports success but the build actually failed

**Cause:** PowerShell's `$ErrorActionPreference = "Stop"` only catches cmdlet errors, not
native command failures (npm, python, pyinstaller, ISCC).

**Fix:** Every external command in build scripts must be followed by:
```powershell
if ($LASTEXITCODE -ne 0) { throw "command failed" }
```

### Vite build injects version as JavaScript expression instead of string

**Cause:** Vite `define` values without `JSON.stringify()` are treated as JS expressions.

**Fix:** Always use `JSON.stringify()` in `vite.config.js`:
```javascript
define: { __APP_VERSION__: JSON.stringify(version) }
```

### Inno Setup ignores /D version override

**Cause:** An unconditional `#define` in the `.iss` file overrides the command-line `/D` flag.

**Fix:** Use `#ifndef` guard:
```
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif
```

## Runtime Issues

### A startup shortcut remains after uninstalling Virelo

The machine-wide uninstaller deliberately does not modify any account's per-user files. Before
uninstalling, turn off **Run at Startup** from Virelo's tray menu in each account that enabled it.
If Virelo is already uninstalled, remove `Virelo.lnk` from
`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup` for that account. Per-user settings,
logs, and folder-view recovery backups are retained.

### UAC prompt appears every time the app starts

**Expected behavior.** Virelo requires administrator privileges for global keyboard hooks and
cross-process window manipulation. The app auto-elevates via `ShellExecuteW("runas", ...)`.
Approve UAC with the same Windows account. Supplying another administrator account's credentials
causes Windows to run Virelo under that account, so per-user settings, backups, and Explorer state
would belong to the credential account; that scenario is not supported.

### Window cannot be dragged

**Cause:** The frameless window uses `WM_NCHITTEST` to define drag zones. If the title bar
height constant does not match the frontend, dragging may not work.

**Fix:** Ensure `TITLE_BAR_HEIGHT` in `virelo/app/window.py` matches the frontend TitleBar
component height (34px content + 1px border = 35px).

### Window resize breaks on multi-monitor setup

**Cause:** Monitors to the left of or above the primary monitor have negative screen
coordinates. The `WM_NCHITTEST` handler must use signed 16-bit extraction for `lParam`.

**Fix:** Coordinates must use `ctypes.c_short(val).value` for signed decoding, not
unsigned `val & 0xFFFF`.

### Explorer column auto-size not working

**Cause:** Explorer auto-size uses COM automation and `IColumnManager`, which requires the
target Explorer tab to be in Details view. Explorer may also be temporarily unavailable
during navigation or a shell restart.

**Fix:** Switch the Explorer window to Details view through the View menu or press
`Ctrl+Shift+6`. Leave the folder open briefly so the view can settle. If Explorer was just
restarted, reopen the folder and try again.

### Explorer auto-size uses unexpected CPU while windows are minimized

**Cause:** Older builds could keep an overdue retry pending and poll Explorer every 5
milliseconds while a target window was not interactive.

**Fix:** Install or build the audited version. It delays retries for noninteractive windows
and progressively backs off to a one-second poll while idle. If the problem persists, set
`VIRELO_DEBUG=1`, reproduce it briefly, and inspect `%LOCALAPPDATA%\Virelo\virelo.log`.

### Restoring folder views after applying Details defaults

The Details-default action intentionally clears existing per-folder view customizations so
stale ShellBags cannot override the new defaults. It creates a verified backup first.

Open Virelo's Explorer page and select **Restore latest backup**. Recovery deletes the
current affected keys before writing the fixed-key snapshot, so keys that were originally absent
are restored as absent. Virelo also creates a pre-restore safety backup. If in-app recovery
fails, preserve the reported directory under `%LOCALAPPDATA%\Virelo\view-backup-*` and do
not run another folder-view action until the error is diagnosed.

Do not apply, reset, or restore folder views while Explorer is copying, moving, renaming,
or deleting files. Virelo does not close Explorer automatically. Restart File Explorer or sign
out after the action so Windows reloads the changed folder-view state.

## Smoke Test

### Smoke test fails on QWebEngine construction

**Cause:** QWebEngine may require specific runtime libraries. In a PyInstaller bundle, the
Qt WebEngine process binary must be present.

**Fix:** Verify `dist/Virelo/` contains the QtWebEngineProcess executable and required DLLs.

### Smoke test triggers UAC prompt

**Cause:** The `--smoke-test` flag was not parsed before the admin elevation check.

**Fix:** Ensure `--smoke-test` is parsed before `_is_admin()` in `virelo/app/__main__.py`.
