# Phase 1: Hygiene and Build Pipeline - Pattern Map

**Mapped:** 2026-04-24
**Files analyzed:** 18 (new + modified)
**Analogs found:** 13 / 18 (5 are new files with no existing analog)

## File Classification

| New/Modified File | Role | Data Flow | Closest Analog | Match Quality |
|-------------------|------|-----------|----------------|---------------|
| `app_config.py` | config | transform | `app_config.py` itself | self (extend) |
| `Virelo.spec` | config | transform | `Windows Toolbox.spec` | exact (rename + extend) |
| `installer/virelo.iss` | config | transform | `installer/virelo.iss` itself | self (modify) |
| `frontend/vite.config.js` | config | transform | `frontend/vite.config.js` itself | self (extend) |
| `main.py` | utility | request-response | `main.py` itself | self (remove lines) |
| `frontend/src/app.jsx` | component | request-response | `frontend/src/panels.jsx` | exact (same version pattern) |
| `frontend/src/panels.jsx` | component | request-response | `frontend/src/panels.jsx` itself | self (modify) |
| `frontend/src/pages.jsx` | component | request-response | `frontend/src/pages.jsx` itself | self (modify) |
| `frontend/src/theme.jsx` | utility | transform | `frontend/src/theme.jsx` itself | self (remove comment) |
| `frontend/src/icons.jsx` | utility | transform | `frontend/src/icons.jsx` itself | self (remove comment) |
| `frontend/src/primitives.jsx` | utility | transform | `frontend/src/primitives.jsx` itself | self (remove comment) |
| `scripts/build-installer.ps1` | config | batch | `scripts/build-installer.ps1` itself | self (rewrite) |
| `.gitignore` | config | — | none | no analog |
| `README.md` | config | — | none | no analog |
| `LICENSE` | config | — | none | no analog |
| `CLAUDE.md` | config | — | `CLAUDE.md` (rewrite) | no analog (full rewrite) |
| `scripts/bootstrap.ps1` | config | batch | `scripts/build-installer.ps1` | role-match |
| `scripts/clean.ps1` | config | batch | `scripts/build-installer.ps1` | role-match |
| `scripts/build-frontend.ps1` | config | batch | `scripts/build-installer.ps1` | role-match |
| `scripts/build-app.ps1` | config | batch | `scripts/build-installer.ps1` | role-match |
| `scripts/verify-release.ps1` | config | batch | `scripts/build-installer.ps1` | role-match |

---

## Pattern Assignments

### `app_config.py` (config, extend)

**Analog:** `app_config.py` (self — extend in-place)

**Current state** (lines 1–28 — full file):
```python
APP_NAME = "Virelo"
ORGANIZATION = "Yusuf Qwareeq"
APP_ID = "com.yusufqwareeq.virelo"
LOG_DIR = "Virelo"
LOG_FILE = "virelo.log"
SETTINGS_GROUP = "Settings"

DEFAULTS = {
    "snap_key": "shift",
    ...
}

def normalize_snap_presses(value):
    ...
```

**Pattern to follow:** `UPPER_SNAKE_CASE` module-level constants (established convention per CLAUDE.md). No classes or nested structures — flat key=value at module top.

**What to add** (D-02 full list):
```python
APP_DISPLAY_NAME = "Virelo"
APP_VERSION = "1.5.0"
APP_EXECUTABLE_NAME = "Virelo.exe"
APP_DIST_DIR_NAME = "Virelo"
APP_PUBLISHER = "Yusuf Qwareeq"
APP_SUPPORT_URL = "https://github.com/yusufqwareeq/virelo"
APP_SETTINGS_ORG = "Yusuf Qwareeq"
APP_LOG_DIR = "Virelo"
APP_LOG_FILE = "virelo.log"
```

**Critical:** `main.py` already imports `APP_NAME`, `APP_ID`, `ORGANIZATION`, `LOG_DIR`, `LOG_FILE` by name from `app_config` (lines 33–41). New constants must not break these. `LOG_DIR`/`LOG_FILE` are superseded by `APP_LOG_DIR`/`APP_LOG_FILE` per D-02 — keep old names as aliases or update all imports.

---

### `Virelo.spec` (config, rename + extend)

**Analog:** `Windows Toolbox.spec` (full file — rename and add version extraction)

**Current spec content** (lines 1–51, full file):
```python
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ("icon.ico", "."),
        ("frontend/dist", "frontend/dist"),
    ],
    hiddenimports=[
        'bridge', 'webview', 'settings_state', 'snap_service',
        'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineCore', 'PySide6.QtWebChannel',
    ],
    ...
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='Virelo',
    ...
    upx=True,
    console=False,
    icon=['icon.ico'],
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=True, upx_exclude=[], name='Virelo')
```

**Version extraction pattern to prepend** (from RESEARCH.md Pattern 4):
```python
# -*- mode: python ; coding: utf-8 -*-
import re
from pathlib import Path

# Parse APP_VERSION via regex -- do NOT import app_config directly.
# Importing app_config in spec context may trigger PySide6 import chain.
_cfg = Path("app_config.py").read_text()
_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', _cfg)
APP_VERSION = _match.group(1) if _match else "0.0.0"
```

**Critical anti-pattern:** Never use `from app_config import APP_VERSION` in the spec file. Use regex text parsing instead (RESEARCH.md Pitfall 4).

---

### `installer/virelo.iss` (config, modify)

**Analog:** `installer/virelo.iss` (self — modify version define at line 5)

**Current state** (lines 1–5):
```pascal
#define MyAppName "Virelo"
#define MyAppPublisher "Yusuf Qwareeq"
#define MyAppURL "mailto:qwareeq8@gmail.com"
#define MyAppExeName "Virelo.exe"
#define MyAppVersion "1.4.2"
```

**Target state** (replace line 5 with `#ifndef` guard):
```pascal
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif
```

**Critical:** The `#define MyAppVersion "1.4.2"` must become `#ifndef` — an unconditional `#define` silently overrides the `/D` flag passed by the build script (RESEARCH.md Pitfall 3).

---

### `frontend/vite.config.js` (config, extend)

**Analog:** `frontend/vite.config.js` (self — add `define` block)

**Current state** (lines 1–20, full file):
```javascript
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    outDir: 'dist',
    assetsInlineLimit: 100000,
    rollupOptions: {
      output: {
        manualChunks: undefined,
      },
    },
  },
  server: {
    port: 5173,
    strictPort: true,
  },
});
```

**`define` block to insert** after `base: './'`:
```javascript
define: {
  __APP_VERSION__: JSON.stringify(process.env.VITE_APP_VERSION || 'dev'),
},
```

**Critical:** Must use `JSON.stringify()`. Without it, `process.env.VITE_APP_VERSION` is treated as a JS expression, not a string (RESEARCH.md Pitfall 1).

---

### `main.py` (utility, remove lines)

**Analog:** `main.py` (self — targeted removals only)

**Lines to remove — deprecated Qt attributes** (lines 1420–1421):
```python
# REMOVE these two lines:
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps)
```
Context: These appear inside the `main()` function between `_enable_dpi_awareness()` call (line 1414) and `setOrganizationName` (line 1422). Remove both lines; surrounding code is unchanged.

**Migration-phase comments to clean** (edit-in-place, do not remove the surrounding logic):

| Line | Current content | Action |
|------|-----------------|--------|
| 908 | `Phase 7: UI replaced with React frontend...` | Rewrite docstring without "Phase 7" label |
| 958 | `# Frameless + resizable window (SC-6).` | Remove `(SC-6)` suffix |
| 991 | `# --- Bridge + WebView (Phase 7) ---` | Remove `(Phase 7)` suffix |
| 1353 | `# Resizable window via WM_NCHITTEST (SC-6, IC-11)` | Remove `(SC-6, IC-11)` suffix |

**Import pattern** (lines 33–41 — add new constants after cleanup):
```python
from app_config import (
    APP_ID,
    APP_NAME,
    DEFAULTS,
    LOG_DIR,
    LOG_FILE,
    ORGANIZATION,
    normalize_snap_presses,
)
```
When new `app_config.py` constants are added, update this import block to include any renamed constants (e.g., `APP_LOG_DIR` replacing `LOG_DIR`).

---

### `frontend/src/app.jsx` (component, modify)

**Analog:** `frontend/src/app.jsx` (self — line 105)

**Current** (line 105):
```jsx
v1.4.2 · up to date
```

**Target:**
```jsx
v{__APP_VERSION__} · up to date
```

**Pattern:** `__APP_VERSION__` is a global injected by Vite `define`. No import needed. Use directly in JSX as a JS expression `{__APP_VERSION__}`.

---

### `frontend/src/panels.jsx` (component, modify)

**Analog:** `frontend/src/panels.jsx` (self)

**Line 1 — v2 comment to remove:**
```jsx
// Command palette (Ctrl/Cmd+K) for Virelo v2.
```
Remove this entire line. Replacement: either no comment or a plain-language description without "v2".

**Line 127 — version string to replace:**
```jsx
// Current:
<span>Virelo 1.4.2</span>

// Target:
<span>Virelo {__APP_VERSION__}</span>
```

---

### `frontend/src/pages.jsx` (component, modify)

**Analog:** `frontend/src/pages.jsx` (self)

**Line 1 — v2 comment to remove:**
```jsx
// Pages for v2: Window Snap, Explorer, Shortcuts, General, About.
```

**Line 249 — version string to replace:**
```jsx
// Current:
<span>Version 1.4.2</span>

// Target:
<span>Version {__APP_VERSION__}</span>
```

**Line 232 — do NOT change:**
```jsx
{ v: '1.4.2', date: 'Apr 12', items: [...] }
```
This is historical changelog data, not a current-version display (RESEARCH.md Stale Reference Inventory note).

---

### `frontend/src/theme.jsx` (utility, modify)

**Line 1 — v2 comment to remove:**
```jsx
// Theme tokens + context for Virelo v2.
```
Remove the `v2` label. Keep the rest of the comment or replace with a clean description.

---

### `frontend/src/icons.jsx` (utility, modify)

**Line 1 — v2 comment to remove:**
```jsx
// SVG icon component extracted from v2/primitives.jsx.
```
Remove the `v2` label.

---

### `frontend/src/primitives.jsx` (utility, modify)

**Line 1 — v2 comment to remove:**
```jsx
// Shared UI primitives for v2.
```
Remove the `v2` label.

---

### `scripts/build-installer.ps1` (config/batch, rewrite)

**Analog:** `scripts/build-installer.ps1` (self — rewrite based on Pattern 3 from RESEARCH.md)

**Current state** (lines 1–27, full file):
```powershell
$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

if (-not (Test-Path "Windows Toolbox.spec")) {
    throw "PyInstaller spec not found: Windows Toolbox.spec"
}

Write-Host "Building EXE with PyInstaller..."
& "$projectRoot\.venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm "Windows Toolbox.spec"

$IsccPath = $env:ISCC_PATH
if (-not $IsccPath) {
    $candidates = @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $IsccPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

if (-not $IsccPath) {
    throw "ISCC.exe not found. Install Inno Setup 6 or set ISCC_PATH."
}

Write-Host "Building installer with Inno Setup..."
& $IsccPath "installer\virelo.iss"
```

**Rewrite follows three established patterns from the current file:**
1. `$ErrorActionPreference = "Stop"` at top — keep
2. `Resolve-Path "$PSScriptRoot\.."` for project root — keep
3. ISCC.exe candidate search (`$candidates` array) — keep

**New pattern to add** (version extraction + `$LASTEXITCODE` checks):
```powershell
# Read APP_VERSION from app_config.py
$versionMatch = Select-String -Path "app_config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $versionMatch) { throw "APP_VERSION not found in app_config.py" }
$AppVersion = $versionMatch.Matches.Groups[1].Value
Write-Host "[build-installer] Version: $AppVersion"

# Call build-app.ps1 first (it calls build-frontend.ps1 internally)
& "$PSScriptRoot\build-app.ps1"
if ($LASTEXITCODE -ne 0) { throw "build-app.ps1 failed" }

# Run ISCC with version define
& $IsccPath "/DMyAppVersion=$AppVersion" "installer\virelo.iss"
if ($LASTEXITCODE -ne 0) { throw "ISCC.exe failed with exit code $LASTEXITCODE" }
```

**Critical:** Every external command call (`npm`, `python`, `pyinstaller`, ISCC) must be followed by `if ($LASTEXITCODE -ne 0) { throw "..." }` (RESEARCH.md Pitfall 2).

---

### `scripts/bootstrap.ps1` (config/batch, new)

**Analog:** `scripts/build-installer.ps1` (role-match — same PowerShell build script pattern)

**Pattern to copy** (boilerplate from build-installer.ps1 lines 1–4):
```powershell
$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot
```

**Core behavior per RESEARCH.md:**
```powershell
# Create .venv if absent
if (-not (Test-Path ".venv")) {
    Write-Host "[bootstrap] Creating .venv..."
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "python -m venv failed" }
}

# Install Python dependencies
Write-Host "[bootstrap] Installing Python dependencies..."
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
```

---

### `scripts/clean.ps1` (config/batch, new)

**Analog:** `scripts/build-installer.ps1` (role-match)

**Boilerplate pattern** (same as bootstrap.ps1):
```powershell
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot
```

**Core behavior:** Remove `build/`, `dist/`, `frontend/dist/`, `__pycache__/`, `.ruff_cache/`, `.pytest_cache/`, `*.spec.bak`.

---

### `scripts/build-frontend.ps1` (config/batch, new)

**Analog:** `scripts/build-installer.ps1` (role-match — best match for pattern)

**Full pattern** (from RESEARCH.md Pattern 3 — build-frontend example):
```powershell
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

# --- Precondition checks ---
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) { throw "Node.js not found. Install from https://nodejs.org" }

$npm = Get-Command npm -ErrorAction SilentlyContinue
if (-not $npm) { throw "npm not found. Install Node.js from https://nodejs.org" }

Write-Host "[build-frontend] Node $(node --version), npm $(npm --version)"

# --- Read version for VITE_APP_VERSION ---
$versionMatch = Select-String -Path "app_config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $versionMatch) { throw "APP_VERSION not found in app_config.py" }
$env:VITE_APP_VERSION = $versionMatch.Matches.Groups[1].Value
Write-Host "[build-frontend] Version: $env:VITE_APP_VERSION"

# --- Build ---
Push-Location frontend
if (-not (Test-Path "node_modules")) {
    Write-Host "[build-frontend] Installing dependencies..."
    npm ci
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed" }
}

npm run build
if ($LASTEXITCODE -ne 0) { throw "npm run build failed" }
Pop-Location

# --- Postcondition check ---
if (-not (Test-Path "frontend\dist\index.html")) {
    throw "Frontend build failed: frontend\dist\index.html not found"
}
Write-Host "[build-frontend] OK: frontend/dist/index.html exists"
```

---

### `scripts/build-app.ps1` (config/batch, new)

**Analog:** `scripts/build-installer.ps1` (role-match)

**Pattern:**
```powershell
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

# --- Precondition checks ---
$python = Get-Command ".venv\Scripts\python.exe" -ErrorAction SilentlyContinue
if (-not $python) { throw ".venv not found. Run bootstrap.ps1 first." }

# --- Call build-frontend.ps1 first ---
& "$PSScriptRoot\build-frontend.ps1"
if ($LASTEXITCODE -ne 0) { throw "build-frontend.ps1 failed" }

# --- Run PyInstaller ---
Write-Host "[build-app] Running PyInstaller..."
& ".venv\Scripts\python.exe" -m PyInstaller --clean --noconfirm "Virelo.spec"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

# --- Postcondition check ---
if (-not (Test-Path "dist\Virelo\Virelo.exe")) {
    throw "Build failed: dist\Virelo\Virelo.exe not found"
}
Write-Host "[build-app] OK: dist/Virelo/Virelo.exe exists"
```

---

### `scripts/verify-release.ps1` (config/batch, new)

**Analog:** `scripts/build-installer.ps1` (role-match)

**Pattern:** Boilerplate + Test-Path checks for all expected artifacts without running any build steps.

---

## Shared Patterns

### PowerShell Boilerplate (All `scripts/*.ps1`)
**Source:** `scripts/build-installer.ps1` lines 1–4
```powershell
$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot
```
**Apply to:** Every PowerShell script in `scripts/`.

### PowerShell Version Extraction
**Source:** RESEARCH.md Pattern 5 (no existing analog in codebase yet)
```powershell
$versionMatch = Select-String -Path "app_config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $versionMatch) { throw "APP_VERSION not found in app_config.py" }
$AppVersion = $versionMatch.Matches.Groups[1].Value
```
**Apply to:** `build-frontend.ps1`, `build-app.ps1`, `build-installer.ps1`.

### PowerShell External Command Guard
**Source:** RESEARCH.md Pitfall 2 / Pattern 3 (no existing analog — current script omits this)
```powershell
& some-external-command
if ($LASTEXITCODE -ne 0) { throw "command failed with exit code $LASTEXITCODE" }
```
**Apply to:** Every call to `npm`, `python`, `pyinstaller`, `ISCC.exe` in all build scripts.

### ISCC.exe Discovery
**Source:** `scripts/build-installer.ps1` lines 13–20
```powershell
$IsccPath = $env:ISCC_PATH
if (-not $IsccPath) {
    $candidates = @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    $IsccPath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $IsccPath) {
    throw "ISCC.exe not found. Install Inno Setup 6 from https://jrsoftware.org/isinfo.php or set ISCC_PATH."
}
```
**Apply to:** `build-installer.ps1` only.

### Vite `define` Version Injection
**Source:** RESEARCH.md Pattern 1 / `frontend/vite.config.js` (extend)
```javascript
define: {
  __APP_VERSION__: JSON.stringify(process.env.VITE_APP_VERSION || 'dev'),
},
```
**Apply to:** `frontend/vite.config.js` only (all JSX files consume `__APP_VERSION__` as a global).

### JSX Version Display
**Source:** `frontend/src/panels.jsx` line 127 (after modification)
```jsx
{__APP_VERSION__}
```
**Apply to:** `frontend/src/app.jsx` line 105, `frontend/src/panels.jsx` line 127, `frontend/src/pages.jsx` line 249.

### Python Config Constant Style
**Source:** `app_config.py` lines 1–6
```python
APP_NAME = "Virelo"
ORGANIZATION = "Yusuf Qwareeq"
APP_ID = "com.yusufqwareeq.virelo"
```
**Apply to:** All new constants added to `app_config.py` — `UPPER_SNAKE_CASE`, module-level, string literals, no classes.

---

## No Analog Found

Files with no close match in the codebase (planner should use RESEARCH.md patterns and standard templates):

| File | Role | Data Flow | Reason |
|------|------|-----------|--------|
| `.gitignore` | config | — | No .gitignore exists. Use GitHub Python template + project additions from RESEARCH.md `.gitignore Pattern Reference` section. |
| `README.md` | config | — | No README exists. D-11 specifies all required sections. |
| `LICENSE` | config | — | No LICENSE exists. Use standard SPDX MIT text. |
| `CLAUDE.md` | config | — | Existing CLAUDE.md is auto-generated GSD content with no reusable structure. D-13 specifies content; write fresh. |

---

## Metadata

**Analog search scope:** Project root, `scripts/`, `frontend/src/`, `installer/`, `app_config.py`, `main.py`, `Windows Toolbox.spec`
**Files scanned:** 13 source files read directly
**Pattern extraction date:** 2026-04-24

**Key constraint confirmed:** `main.py` imports `APP_NAME`, `APP_ID`, `ORGANIZATION`, `LOG_DIR`, `LOG_FILE` from `app_config` by name. When adding the D-02 constants (`APP_LOG_DIR`, `APP_LOG_FILE`), the planner must decide whether to rename `LOG_DIR`/`LOG_FILE` (requires updating `main.py` imports) or keep them as-is and add new names alongside.
