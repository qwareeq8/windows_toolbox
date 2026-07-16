# Virelo

A personal Windows desktop utility that centers windows, auto-sizes File Explorer columns, and applies consistent Explorer Details-view defaults.

## What It Does

Virelo provides three core features:

1. **Window Snap.** A multi-press keyboard shortcut resizes the foreground window and centers its visible frame on the current monitor. Fixed-size windows retain their size and are only centered. Hold the configured Restore Key during the shortcut to restore the window to its original dimensions.

2. **Explorer Column Auto-Size.** Automatically fits every visible column when navigating between folders in Details view.

3. **Explorer Details Defaults.** Applies Details as the default for all Explorer folder types using the same Windows registry model as [WinSetView](https://github.com/LesFerch/WinSetView), with a single opinionated action instead of a full configuration surface. Virelo creates a verified recovery backup before clearing existing folder customizations, can reset to Windows defaults, and can restore the latest backup.

Do not start a folder-view action while Explorer is copying, moving, renaming, or deleting files. Virelo does not close Explorer automatically. After the action, manually restart File Explorer or sign out so Windows reloads the changed state. Backups are stored under `%LOCALAPPDATA%\Virelo\view-backup-*` and include a recovery manifest.

## Requirements

- **Windows 10 version 1809 (build 17763) or later, or Windows 11, on x64.** Virelo uses Win32 APIs, COM automation, and Windows-specific system features.
- **Administrator privileges.** The application auto-elevates at launch through a UAC prompt. Elevation allows global keyboard hooks and window manipulation across process integrity levels.
- **Personal-use software.** Virelo has no telemetry, accounts, or automatic updates.

## Build from Source

### Prerequisites

- Install Git.
- Install 64-bit Python 3.12 to 3.14.
- Install Node.js 22 to 24 for the frontend build.
- Install Inno Setup 6 only when building the installer.

### Build Steps

```powershell
scripts/bootstrap.ps1 -Recreate  # Create the locked development environment.
scripts/build-app.ps1             # Build the one-folder application.
# Or, for the complete installer pipeline:
scripts/build-installer.ps1
```

Each script validates its preconditions and fails early with a clear error message if a required tool is missing.

## Development Mode

Set the `VIRELO_DEV` environment variable, start the Vite dev server, then run the app:

```powershell
$env:VIRELO_DEV = 1
cd frontend
npm run dev
```

In a separate terminal:

```powershell
.venv\Scripts\python.exe main.py
```

The frontend hot-reloads from `localhost:5173`. Changes to React components appear immediately without restarting the Python backend.

## Architecture

Virelo uses a Python/PySide6 backend that hosts a React frontend inside a QWebEngineView. The two layers communicate through QWebChannel:

- **Python backend.** Owns all OS-level logic: window management, keyboard hooks, COM automation, system tray, and settings persistence through QSettings.
- **React frontend.** Renders the settings UI, theme controls, Explorer recovery actions, and command palette inside the embedded Chromium browser.
- **QWebChannel bridge.** A single `VireloBridge` QObject is the communication boundary. Data crosses the bridge as JSON strings.

In release mode, the frontend is built to static files (`frontend/dist/`) and loaded via `file://` URL. In dev mode, it connects to the Vite dev server.

## Documentation

- [Building from Source](docs/BUILD.md). Full build pipeline, development mode, and version management.
- [Troubleshooting](docs/TROUBLESHOOTING.md). Common build and runtime issues with fixes.
- [Release Checklist](docs/RELEASE.md). Step-by-step release and acceptance process.
- [Verification and Audit, 2026-07-16](docs/AUDIT-2026-07-16.md). Findings, fixes, evidence, and explicit verification boundaries.

## Status

Under active development. The current version is defined by `APP_VERSION` in `virelo/app/config.py`.

## License

[MIT](LICENSE)
