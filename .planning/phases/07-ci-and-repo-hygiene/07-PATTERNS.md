# Phase 7: CI and Repo Hygiene - Pattern Map

**Mapped:** 2026-04-24
**Files analyzed:** 11 modified files (0 new files)
**Analogs found:** 11 / 11

---

## File Classification

| Modified File | Role | Data Flow | Closest Analog | Match Quality |
|---------------|------|-----------|----------------|---------------|
| `.github/workflows/ci.yml` | config (CI) | event-driven | `.github/workflows/ci.yml` (self) | exact — extend existing jobs |
| `scripts/clean.ps1` | utility (build) | batch | `scripts/clean.ps1` (self) | exact — extend existing targets |
| `virelo/services/snap.py` | service | event-driven | `virelo/services/snap.py` (self) | exact — rename two class definitions |
| `virelo/app/window.py` | config (app shell) | event-driven | `virelo/app/window.py` (self) | exact — rename import + two instantiations |
| `tests/unit/test_snap_geometry.py` | test | request-response | `tests/unit/test_snap_geometry.py` (self) | exact — add skipif marker + rename import |
| `tests/unit/test_app_config.py` | test | request-response | `tests/unit/test_app_config.py` (self) | exact — rephrase one docstring |
| `frontend/package.json` | config | — | `virelo/app/config.py` | role-match — version string must match config.py |
| `installer/virelo.iss` | config (installer) | — | `virelo/app/config.py` | role-match — URL must match APP_SUPPORT_URL |
| `README.md` | docs | — | `pyproject.toml` | role-match — Python version must match requires-python |
| `CLAUDE.md` | docs | — | `virelo/services/snap.py` | role-match — class names must match snap.py |

---

## Pattern Assignments

### `.github/workflows/ci.yml` (config/CI, event-driven)

**Analog:** `.github/workflows/ci.yml` — the file itself; two independent modifications.

**Existing job structure** (lines 1-66, full file):
```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
          cache: 'pip'
      - run: pip install ruff
      - run: ruff check .
      - run: ruff format --check .

  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'
          cache: 'pip'
      - run: pip install -e ".[dev]"
      - run: pytest tests/unit/ -q

  frontend:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '22'
          cache: 'npm'
          cache-dependency-path: frontend/package-lock.json
      - run: npm ci
        working-directory: frontend
      - run: npm run build
        working-directory: frontend
      - run: npx vitest run
        working-directory: frontend

  stale-name:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Check for stale "Windows Toolbox" references
        run: |
          if grep -rn "Windows Toolbox" . \
            --include="*.py" --include="*.jsx" --include="*.js" \
            --include="*.json" --include="*.toml" --include="*.yml" \
            --include="*.iss" --include="*.spec" \
            --exclude-dir=node_modules --exclude-dir=.git \
            --exclude-dir=dist --exclude-dir=build \
            --exclude-dir=.planning; then
            echo "ERROR: Found stale 'Windows Toolbox' references"
            exit 1
          fi
          echo "OK: No stale references found"
```

**Change 1 — stale-name grep fix** (line 61): add `--exclude-dir=.github` to the grep flags. Current line reads:
```yaml
            --exclude-dir=.planning; then
```
Target state:
```yaml
            --exclude-dir=.planning --exclude-dir=.github; then
```

**Change 2 — new version-check job**: append after the `stale-name` job, following the exact same `runs-on: ubuntu-latest` + `actions/checkout@v4` structure as all other jobs. No `actions/setup-python` needed — only `grep` (ubuntu built-in):
```yaml
  version-check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Check version consistency (config.py vs package.json)
        run: |
          PY_VERSION=$(grep -oP '(?<=APP_VERSION = ")[^"]+' virelo/app/config.py)
          PKG_VERSION=$(grep -oP '(?<="version": ")[^"]+' frontend/package.json)
          echo "config.py APP_VERSION:  $PY_VERSION"
          echo "package.json version:   $PKG_VERSION"
          if [ "$PY_VERSION" != "$PKG_VERSION" ]; then
            echo "ERROR: Version mismatch"
            exit 1
          fi
          echo "OK: Versions match ($PY_VERSION)"
```

---

### `scripts/clean.ps1` (utility/build, batch)

**Analog:** `scripts/clean.ps1` — the file itself; extend, do not rewrite.

**Existing file** (lines 1-29, full file):
```powershell
$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

$targets = @(
    "build",
    "dist",
    "frontend\dist",
    "__pycache__",
    ".pytest_cache",
    ".ruff_cache"
)

foreach ($target in $targets) {
    if (Test-Path $target) {
        Write-Host "[clean] Removing $target"
        Remove-Item -Recurse -Force $target
    }
}

# Clean *.spec.bak files
Get-ChildItem -Filter "*.spec.bak" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "[clean] Removing $($_.Name)"
    Remove-Item -Force $_.FullName
}

Write-Host "[clean] OK: Build artifacts removed"
```

**Change 1 — add `installer\dist` to `$targets` array** (line 10): Use backslash to match the existing `frontend\dist` style:
```powershell
$targets = @(
    "build",
    "dist",
    "frontend\dist",
    "installer\dist",    # NEW (D-04)
    ".pytest_cache",
    ".ruff_cache"
)
```
Note: Remove `"__pycache__"` from the flat `$targets` array — it is replaced by the recursive pass below (D-03). Keeping it would only clean the root-level `__pycache__` and fail silently on subdirectory caches.

**Change 2 — recursive `__pycache__` and `*.pyc` removal**: insert between the `foreach` loop and the `Write-Host "[clean] OK"` line. Copy the `Get-ChildItem -Filter "*.spec.bak"` pattern already in the file:
```powershell
# Recursive __pycache__ and *.pyc removal (D-03)
Get-ChildItem -Recurse -Filter "__pycache__" -Directory -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host "[clean] Removing $($_.FullName)"
        Remove-Item -Recurse -Force $_.FullName
    }

Get-ChildItem -Recurse -Filter "*.pyc" -ErrorAction SilentlyContinue |
    ForEach-Object {
        Remove-Item -Force $_.FullName
    }
```

---

### `virelo/services/snap.py` (service, event-driven)

**Analog:** `virelo/services/snap.py` — the file itself; two class definition renames.

**Module docstring** (lines 1-7) — update both old names:
```python
"""Snap service, HotkeyListener, and ShiftSnapRestore engine.

HotkeyListener detects multi-press keyboard patterns and emits a trigger signal.
ShiftSnapRestore performs window snap and restore operations.
...
"""
```
Target state — every mention of old names in the docstring updated:
```python
"""Snap service, MultiPressHotkeyListener, and SnapRestoreController engine.

MultiPressHotkeyListener detects multi-press keyboard patterns and emits a trigger signal.
SnapRestoreController performs window snap and restore operations.
...
"""
```

**HotkeyListener class definition** (line 49):
```python
class HotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""
```
Target state:
```python
class MultiPressHotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""
```

**ShiftSnapRestore class definition** (line 134):
```python
class ShiftSnapRestore(QtCore.QObject):
    """Performs window snap and restore operations."""
```
Target state:
```python
class SnapRestoreController(QtCore.QObject):
    """Performs window snap and restore operations."""
```

**LOG message in ShiftSnapRestore.perform** (line 208) — old name appears in exception message:
```python
            LOG.exception("ShiftSnapRestore.perform failed.", exc_info=e)
```
Target state:
```python
            LOG.exception("SnapRestoreController.perform failed.", exc_info=e)
```

**SnapService docstrings** (lines 329 and 334) — old names appear in docstrings:
```python
    def __init__(self, shift_mgr):
        """Accept a ShiftSnapRestore instance (or None during early init)."""
```
```python
    def set_manager(self, mgr):
        """Set or replace the ShiftSnapRestore instance."""
```
Target state:
```python
    def __init__(self, shift_mgr):
        """Accept a SnapRestoreController instance (or None during early init)."""
```
```python
    def set_manager(self, mgr):
        """Set or replace the SnapRestoreController instance."""
```

**SnapService.set_listener docstring** (line 338):
```python
    def set_listener(self, listener):
        """Set or replace the HotkeyListener instance."""
```
Target state:
```python
    def set_listener(self, listener):
        """Set or replace the MultiPressHotkeyListener instance."""
```

---

### `virelo/app/window.py` (app shell, event-driven)

**Analog:** `virelo/app/window.py` — the file itself; rename import and two instantiation sites.

**Import line** (line 23):
```python
from virelo.services.snap import HotkeyListener, ShiftSnapRestore, SnapService
```
Target state:
```python
from virelo.services.snap import MultiPressHotkeyListener, SnapRestoreController, SnapService
```

**Comment at line 174** (references old class name in comment text):
```python
        # snap_enabled used by business logic (ShiftSnapRestore, _test_snap)
```
Target state:
```python
        # snap_enabled used by business logic (SnapRestoreController, _test_snap)
```

**Comment at line 197**:
```python
        # HotkeyListener + ShiftSnapRestore (per D-01/D-02/D-03)
```
Target state:
```python
        # MultiPressHotkeyListener + SnapRestoreController (per D-01/D-02/D-03)
```

**Instantiation at line 198** — class name only; variable name `_hotkey_listener` is unchanged per D-14:
```python
        self._hotkey_listener = HotkeyListener(self.settings)
```
Target state:
```python
        self._hotkey_listener = MultiPressHotkeyListener(self.settings)
```

**Instantiation at line 199** — class name only; variable name `shift_mgr` is unchanged per D-13:
```python
        self.shift_mgr = ShiftSnapRestore(self.settings)
```
Target state:
```python
        self.shift_mgr = SnapRestoreController(self.settings)
```

---

### `tests/unit/test_snap_geometry.py` (test, request-response)

**Analog:** `tests/unit/test_snap_geometry.py` — the file itself; two changes.

**Change 1 — add `sys` import and `@pytest.mark.skipif` to `test_restore_maximized_window`**.

No `sys` import exists in this file currently. Add it at module level alongside the existing imports. The conftest stubs `win32con`/`win32gui` at module level, but `ctypes.wintypes` inside the test body is not stubbed and raises `ImportError` on Linux.

Current module-level imports (lines 1-11):
```python
"""Tests for snap geometry and fullscreen detection (QUAL-03).
...
"""

from unittest.mock import MagicMock, patch

from virelo.app.config import DEFAULTS
from virelo.platform.win32_helpers import FULLSCREEN_TOLERANCE, _rect_matches_monitor
from virelo.services.snap import calculate_snap_position
```
Add `import sys` and `import pytest` (pytest is already an implicit dependency but must be imported for the decorator). Insert after the existing imports:
```python
import sys

import pytest
```

Current function signature (line 108):
```python
def test_restore_maximized_window():
    """Restore of a previously-maximized window issues SW_MAXIMIZE (SNAP-04)."""
```
Target state — decorator placed immediately before `def`, no blank line between decorator and function:
```python
@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs only available on Windows")
def test_restore_maximized_window():
    """Restore of a previously-maximized window issues SW_MAXIMIZE (SNAP-04)."""
```

**Change 2 — rename `ShiftSnapRestore` import and usage inside the test body**.

Current import at line 116:
```python
    from virelo.services.snap import ShiftSnapRestore
```
Target state:
```python
    from virelo.services.snap import SnapRestoreController
```

Current instantiation at line 119:
```python
    mgr = ShiftSnapRestore.__new__(ShiftSnapRestore)
```
Target state:
```python
    mgr = SnapRestoreController.__new__(SnapRestoreController)
```

Current section comment at line 105:
```python
# -- ShiftSnapRestore restore --
```
Target state:
```python
# -- SnapRestoreController restore --
```

Current docstring at line 118:
```python
    # Create ShiftSnapRestore bypassing __init__ (avoids keyboard hooks and EnumWindows)
```
Target state:
```python
    # Create SnapRestoreController bypassing __init__ (avoids keyboard hooks and EnumWindows)
```

---

### `tests/unit/test_app_config.py` (test, request-response)

**Analog:** `tests/unit/test_app_config.py` — the file itself; rephrase one docstring.

**Current docstring at line 49** — contains the literal banned string "Windows Toolbox":
```python
def test_app_name_is_virelo():
    """APP_NAME must be 'Virelo' (never 'Windows Toolbox')."""
    assert APP_NAME == "Virelo"
```
Target state — indirect reference avoids triggering the stale-name grep:
```python
def test_app_name_is_virelo():
    """APP_NAME must be 'Virelo' (not the old product name)."""
    assert APP_NAME == "Virelo"
```

---

### `frontend/package.json` (config, version sync)

**Version source analog:** `virelo/app/config.py` line 4 — `APP_VERSION = "1.5.0"`.

**Current state** (line 4):
```json
  "version": "1.4.2",
```
Target state — must match `APP_VERSION` in config.py:
```json
  "version": "1.5.0",
```
All other fields remain unchanged.

---

### `installer/virelo.iss` (config/installer)

**URL source analog:** `virelo/app/config.py` line 8 — `APP_SUPPORT_URL = "https://github.com/yusufqwareeq/virelo"`.

**Current state** (line 3):
```iss
#define MyAppURL "mailto:qwareeq8@gmail.com"
```
Target state — `MyAppURL` has no `#ifndef` guard (unlike `MyAppVersion`) so a direct `#define` edit is safe:
```iss
#define MyAppURL "https://github.com/yusufqwareeq/virelo"
```

---

### `README.md` (docs)

**Version source analog:** `pyproject.toml` `requires-python = ">=3.12"`.

**Current state** (line 24):
```markdown
- Python 3.8+
```
Target state:
```markdown
- Python 3.12+
```

---

### `CLAUDE.md` (docs)

**Class name source analog:** `virelo/services/snap.py` lines 49 and 134 (after rename).

**Current state** (line 56 of CLAUDE.md):
```
    - `snap.py` -- HotkeyListener (keyboard detection), ShiftSnapRestore (window movement), SnapService facade, geometry calculations
```
Target state:
```
    - `snap.py` -- MultiPressHotkeyListener (keyboard detection), SnapRestoreController (window movement), SnapService facade, geometry calculations
```

---

## Shared Patterns

### PowerShell `$ErrorActionPreference` + cmdlet loop
**Source:** `scripts/clean.ps1` lines 1-20
**Apply to:** All PowerShell script edits in this phase
```powershell
$ErrorActionPreference = "Stop"
# ...
foreach ($target in $targets) {
    if (Test-Path $target) {
        Write-Host "[clean] Removing $target"
        Remove-Item -Recurse -Force $target
    }
}
```
Existing recursive file pattern in the same file (lines 23-26) — copy this `Get-ChildItem | ForEach-Object` style for the `__pycache__` and `*.pyc` passes:
```powershell
Get-ChildItem -Filter "*.spec.bak" -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Host "[clean] Removing $($_.Name)"
    Remove-Item -Force $_.FullName
}
```

### CI job skeleton
**Source:** `.github/workflows/ci.yml` lines 9-20 (lint job)
**Apply to:** New `version-check` job — copy the job-level structure (indentation, `runs-on`, `steps`, `uses: actions/checkout@v4`) exactly:
```yaml
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - ...
```

### pytest module-level stub pattern
**Source:** `tests/conftest.py` lines 1-6 — `import sys` at module level is already established in conftest. The `sys.platform` check in `@pytest.mark.skipif` follows the same stdlib-only pattern.
**Apply to:** `test_snap_geometry.py` — add `import sys` + `import pytest` at module level alongside the existing imports.

### No `#ifndef` needed for `MyAppURL`
**Source:** `installer/virelo.iss` lines 1-7
**Apply to:** `MyAppURL` edit only. The `MyAppVersion` constant uses `#ifndef` because it is overridden by `/D` at build time (CLAUDE.md footgun #4). `MyAppURL` is not passed via `/D`, so a plain `#define` replacement is the correct and safe pattern — matching lines 1-2 of virelo.iss.

---

## No Analog Found

All modified files have direct analogs (all are self-modifications). No files in this phase lack a pattern reference.

---

## Metadata

**Analog search scope:** `.github/workflows/`, `scripts/`, `virelo/services/`, `virelo/app/`, `tests/unit/`, `frontend/`, `installer/`, project root docs
**Files read:** 12
**Pattern extraction date:** 2026-04-24
