# Phase 1: Hygiene and Build Pipeline - Research

**Researched:** 2026-04-24
**Domain:** Build pipeline, repository hygiene, version consolidation, PowerShell automation
**Confidence:** HIGH

## Summary

Phase 1 covers three distinct workstreams: (1) identity cleanup -- removing all stale "Windows Toolbox" references and deprecated Qt 6 attributes, (2) repository hygiene -- creating .gitignore, README, LICENSE, and CLAUDE.md, and (3) build pipeline -- consolidating version strings to a single source of truth and creating a multi-script PowerShell build pipeline that produces a working installer from a clean checkout.

The codebase currently has no .gitignore, no README, no LICENSE, no pyproject.toml, and the PyInstaller spec is still named "Windows Toolbox.spec". The version string "1.4.2" is hardcoded in 5 separate locations (installer/virelo.iss, frontend/package.json, panels.jsx, pages.jsx, app.jsx). The build script invokes PyInstaller but never builds the frontend first, meaning a clean checkout cannot produce a working app.

**Primary recommendation:** Define `APP_VERSION` once in `app_config.py`. The build scripts read it via PowerShell regex. Vite injects it at build time via `define`. Inno Setup receives it via ISCC `/D` flag. The PyInstaller spec reads it directly via Python import. Each build script validates its preconditions and fails early.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Version defined once in `app_config.py` as `APP_VERSION`. Frontend receives it via Vite `define` at build time. Inno Setup receives it via `/D` flag from build script. `package.json` version is independent (frontend is private, never published to npm).
- **D-02:** Add all product metadata constants to `app_config.py`: `APP_NAME`, `APP_DISPLAY_NAME`, `APP_VERSION`, `APP_ID`, `APP_EXECUTABLE_NAME`, `APP_DIST_DIR_NAME`, `APP_PUBLISHER`, `APP_SUPPORT_URL`, `APP_SETTINGS_ORG`, `APP_LOG_DIR`, `APP_LOG_FILE`.
- **D-03:** Rename `Windows Toolbox.spec` to `Virelo.spec`. Update all references in `scripts/build-installer.ps1`.
- **D-04:** Remove all internal migration-phase comments (Phase 7, SC-6, IC-11, v2 labels) from production source files.
- **D-05:** Grep verification: `grep -rn "Windows Toolbox" .` and `grep -rn "Toolbox" .` must return zero matches after cleanup.
- **D-06:** Multiple specialized PowerShell scripts in `scripts/`: `bootstrap.ps1`, `clean.ps1`, `build-frontend.ps1`, `build-app.ps1`, `build-installer.ps1`, `verify-release.ps1`. Each script validates preconditions and fails early with clear error messages.
- **D-07:** `build-frontend.ps1` runs `npm ci` if `node_modules` absent, then `npm run build`, then verifies `frontend/dist/index.html` exists.
- **D-08:** `build-app.ps1` calls `build-frontend.ps1` first, then runs PyInstaller with `Virelo.spec`, then verifies `dist/Virelo/Virelo.exe` exists.
- **D-09:** `build-installer.ps1` calls `build-app.ps1` first, locates ISCC.exe, runs `installer/virelo.iss` with version from `app_config.py`, verifies installer output.
- **D-10:** `.gitignore` based on GitHub Python template, extended with: `frontend/node_modules/`, `frontend/dist/`, `build/`, `dist/`, `*.spec.bak`, `.pytest_cache/`, `.ruff_cache/`, `.coverage`, `crash.log`. Do not commit `frontend/dist/` -- treat as build artifact.
- **D-11:** `README.md` covers: what Virelo does, Windows-only requirement, admin privilege requirement, personal-use status, build from source instructions, dev mode instructions, planned features section.
- **D-12:** `LICENSE` file with MIT license.
- **D-13:** `CLAUDE.md` provides build commands, project structure, naming conventions, forbidden changes (no stale names, no fake features, no generated artifacts in git), and known footguns.
- **D-14:** Remove `setAttribute(AA_EnableHighDpiScaling)` and `setAttribute(AA_UseHighDpiPixmaps)` calls from `main.py`. These are no-ops in Qt 6 and produce deprecation warnings.

### Claude's Discretion
- Build script error message formatting and verbosity level
- README structure and section ordering
- CLAUDE.md organization and level of detail
- Whether `bootstrap.ps1` creates a `.venv` or uses the system Python

### Deferred Ideas (OUT OF SCOPE)
None -- discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| IDENT-01 | All "Windows Toolbox" references removed from source, spec files, and build scripts | Grep audit found 3 source locations: `Windows Toolbox.spec` (filename), `build-installer.ps1` (3 references). Also in CLAUDE.md (auto-generated, will be rewritten). D-05 defines verification grep. |
| IDENT-02 | All product metadata defined in one source of truth in app_config.py | `app_config.py` already has APP_NAME, ORGANIZATION, APP_ID. D-02 specifies full list of constants to add. |
| IDENT-03 | Version string flows from app_config.py to installer, frontend, and package metadata | Version "1.4.2" found hardcoded in 5 locations. Vite `define` for frontend, ISCC `/D` for installer, Python import for spec file. |
| IDENT-04 | Internal migration-phase comments removed from production source | Found: `main.py:908` (Phase 7), `main.py:958` (SC-6), `main.py:991` (Phase 7), `main.py:1353` (SC-6, IC-11). Frontend: `theme.jsx:1`, `icons.jsx:1`, `panels.jsx:1`, `pages.jsx:1`, `primitives.jsx:1` (v2 labels). |
| REPO-01 | .gitignore excludes all generated artifacts | No .gitignore exists. Research provides complete pattern list based on GitHub Python template + Node + PyInstaller + Inno Setup additions. |
| REPO-02 | README explains what Virelo does | No README exists. D-11 specifies sections. |
| REPO-03 | LICENSE file present (MIT) | No LICENSE exists. Straightforward file creation. |
| REPO-04 | CLAUDE.md provides Claude Code with build commands and conventions | CLAUDE.md exists but contains auto-generated GSD content. D-13 specifies what it should contain. |
| REPO-05 | Deprecated Qt attributes removed | Lines `main.py:1420-1421` call `setAttribute(AA_EnableHighDpiScaling)` and `setAttribute(AA_UseHighDpiPixmaps)`. Both are no-ops in Qt 6. |
| BUILD-01 | Clean checkout produces working app with one command sequence | Current build script skips frontend build entirely. D-06 through D-09 define the multi-script pipeline. |
| BUILD-02 | Build fails early if required tools are missing | PowerShell precondition validation pattern researched. Check node, npm, python, pyinstaller, ISCC availability. |
| BUILD-03 | Build fails early if frontend/dist is absent after frontend build step | `build-frontend.ps1` must verify `frontend/dist/index.html` exists after `npm run build`. |
| BUILD-04 | PyInstaller spec renamed from "Windows Toolbox.spec" to "Virelo.spec" | Spec file exists at project root with stale name. Rename and update build-installer.ps1 references. |
| BUILD-05 | Inno Setup reads version from generated metadata, not hardcoded string | Currently hardcoded as `#define MyAppVersion "1.4.2"` in virelo.iss. ISCC `/D` flag replaces it. |
| BUILD-06 | Installed app launches and loads React frontend offline | PyInstaller bundles `frontend/dist/` into the frozen app. Build pipeline ensures dist is built first. Manual verification step. |
</phase_requirements>

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Version source of truth | Python backend (`app_config.py`) | -- | Single canonical location; all other systems read from here |
| Frontend version injection | Build tooling (Vite `define`) | -- | Compile-time replacement, no runtime dependency |
| Installer version injection | Build tooling (ISCC `/D` flag) | -- | Preprocessor define passed via command line |
| Build orchestration | Build scripts (PowerShell) | -- | Platform-native scripting, precondition validation |
| Repository metadata | Static files (root directory) | -- | .gitignore, README, LICENSE, CLAUDE.md |
| Deprecated Qt cleanup | Python backend (`main.py`) | -- | Two lines to remove from the `main()` function |
| Stale naming cleanup | All tiers (spec, scripts, source) | -- | Grep-and-replace across the full codebase |

## Standard Stack

### Core (Already Established -- No New Libraries)

This phase adds no new runtime dependencies. It modifies existing files and creates build scripts.

| Tool | Version | Purpose | Why Standard |
|------|---------|---------|--------------|
| PowerShell | 5.1+ (installed) | Build script automation | Already used for `build-installer.ps1`. PowerShell 7.6 also available but 5.1 compatibility ensures scripts work on any Windows 10/11 machine. [VERIFIED: local system check] |
| Vite | ^6.3.4 (installed) | Frontend build + version injection via `define` | Already in project. `define` is the standard mechanism for build-time constants. [VERIFIED: npm registry shows 8.0.10 latest, but project pins ^6.3.4] |
| PyInstaller | >=6.0 (in requirements.txt) | App bundling | Already in project. Spec file is executable Python, can import `app_config` directly. [VERIFIED: requirements.txt] |
| Inno Setup 6 | 6.x (to be installed) | Windows installer | Already in project. ISCC.exe accepts `/D` flag for preprocessor defines. [CITED: jrsoftware.org/ishelp/topic_isppcc.htm] |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `app_config.py` as version source | `pyproject.toml` via `importlib.metadata` | D-01 locks this decision. `app_config.py` is simpler -- no package installation required for the version to be available. `pyproject.toml` version will be added in Phase 3 but `app_config.py` remains the canonical source. |
| PowerShell reading `app_config.py` via regex | Python helper script to emit version | PowerShell regex is simpler (one line) and avoids requiring Python to be available before the first build step. |
| ISCC `/D` flag | `#include` a generated `.iss` fragment | `/D` flag is cleaner -- no temp files, no cleanup needed. [CITED: jrsoftware.org/ishelp/topic_isppcc.htm] |

## Architecture Patterns

### System Architecture: Version Flow

```
app_config.py (APP_VERSION = "1.5.0")
        |
        +---> build-frontend.ps1 ---> reads APP_VERSION via regex
        |         |
        |         +---> passes to: npm run build (via VITE_APP_VERSION env var or vite.config.js define)
        |         |
        |         +---> frontend/dist/ produced with version baked into JS bundle
        |
        +---> build-app.ps1 ---> calls build-frontend.ps1 first
        |         |
        |         +---> PyInstaller reads Virelo.spec
        |                   |
        |                   +---> Virelo.spec imports app_config.APP_VERSION (it's Python)
        |                   |
        |                   +---> dist/Virelo/Virelo.exe produced
        |
        +---> build-installer.ps1 ---> calls build-app.ps1 first
                  |
                  +---> reads APP_VERSION via regex
                  |
                  +---> ISCC.exe /DMyAppVersion=1.5.0 installer/virelo.iss
                  |
                  +---> dist/VireloSetup.exe produced
```

### System Architecture: Build Script Chain

```
bootstrap.ps1          (independent: creates .venv, installs deps)
clean.ps1              (independent: removes build artifacts)

build-frontend.ps1     (preconditions: node, npm)
      ^
      |
build-app.ps1          (preconditions: python, pyinstaller; calls build-frontend.ps1)
      ^
      |
build-installer.ps1    (preconditions: ISCC.exe; calls build-app.ps1)
      ^
      |
verify-release.ps1     (postcondition: checks dist/ output integrity)
```

### Pattern 1: Vite `define` for Build-Time Version Injection

**What:** Vite's `define` config option performs static text replacement at build time. The value is embedded directly into the JavaScript bundle -- no runtime resolution needed.

**When to use:** Injecting constants (version, build date, app name) that the frontend needs at display time but that originate from the backend.

**Example:**
```javascript
// frontend/vite.config.js
// Source: https://vite.dev/config/shared-options.html#define
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  define: {
    __APP_VERSION__: JSON.stringify(process.env.VITE_APP_VERSION || 'dev'),
  },
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

```jsx
// Usage in any React component:
<span>Version {__APP_VERSION__}</span>
```

**Critical detail:** Values passed to `define` must be JSON-serializable. For strings, wrap with `JSON.stringify()` -- otherwise Vite treats the value as a code expression, not a string literal. [CITED: vite.dev/config/shared-options.html#define]

**Alternative approach (also valid):** Instead of reading from `process.env`, the build script can write the version directly into vite.config.js's `define` block. But using an environment variable is cleaner because it avoids modifying vite.config.js during builds.

### Pattern 2: Inno Setup Version via ISCC `/D` Flag

**What:** The ISCC command-line compiler accepts `/D<name>=<value>` to define preprocessor symbols. Combined with `#ifndef` in the .iss file, this provides command-line override with a fallback default.

**When to use:** Passing the version from the build script to the Inno Setup compiler without modifying the .iss file.

**Example:**
```pascal
; installer/virelo.iss
; Source: https://jrsoftware.org/ishelp/topic_isppcc.htm
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif
#define MyAppName "Virelo"
; ... rest of defines ...
```

```powershell
# scripts/build-installer.ps1
$version = (Select-String -Path "app_config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"').Matches.Groups[1].Value
& $IsccPath "/DMyAppVersion=$version" "installer\virelo.iss"
```

**Critical detail:** The `/D` flag with `#ifndef` means the .iss file still works standalone (with the fallback "0.0.0-dev" version) for manual compilation, but the build script overrides it with the real version. [CITED: jrsoftware.org/ishelp/topic_isppcc.htm, jrsoftware.org/ishelp/topic_ifdef.htm]

### Pattern 3: PowerShell Build Script with Precondition Validation

**What:** Each build script starts by validating that its required tools are available, then runs its operations, then validates its output.

**When to use:** Every script in the `scripts/` directory.

**Example:**
```powershell
# scripts/build-frontend.ps1
$ErrorActionPreference = "Stop"
$projectRoot = Resolve-Path "$PSScriptRoot\.."
Set-Location $projectRoot

# --- Precondition checks ---
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) { throw "Node.js not found. Install from https://nodejs.org" }

$npm = Get-Command npm -ErrorAction SilentlyContinue
if (-not $npm) { throw "npm not found. Install Node.js from https://nodejs.org" }

Write-Host "[build-frontend] Node $(node --version), npm $(npm --version)"

# --- Build ---
Push-Location frontend
if (-not (Test-Path "node_modules")) {
    Write-Host "[build-frontend] Installing dependencies..."
    npm ci
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed" }
}

Write-Host "[build-frontend] Building frontend..."
npm run build
if ($LASTEXITCODE -ne 0) { throw "npm run build failed" }
Pop-Location

# --- Postcondition check ---
if (-not (Test-Path "frontend\dist\index.html")) {
    throw "Frontend build failed: frontend\dist\index.html not found"
}

Write-Host "[build-frontend] OK: frontend/dist/index.html exists"
```

**Critical detail:** Use `$ErrorActionPreference = "Stop"` at the top of every script. Check `$LASTEXITCODE` after external commands because `$ErrorActionPreference` does not catch non-zero exit codes from native commands (npm, python, pyinstaller). [ASSUMED]

### Pattern 4: PyInstaller Spec Reading Version from app_config.py

**What:** The spec file is executable Python. It can import modules from the project.

**When to use:** The renamed `Virelo.spec` file.

**Example:**
```python
# Virelo.spec
# -*- mode: python ; coding: utf-8 -*-
import re
from pathlib import Path

# Read version from app_config.py without importing it
# (avoids triggering PySide6 import chain)
_cfg = Path("app_config.py").read_text()
_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', _cfg)
APP_VERSION = _match.group(1) if _match else "0.0.0"

a = Analysis(
    ['main.py'],
    # ... rest of spec ...
)
```

**Critical detail:** Do NOT `import app_config` directly in the spec file. The spec runs in PyInstaller's context where PySide6 may not be importable. Use regex to parse the version string, just like the PowerShell scripts do. [ASSUMED]

### Pattern 5: Reading Python Constants from PowerShell

**What:** PowerShell's `Select-String` with a regex capture group extracts a value from a Python source file.

**When to use:** Any build script that needs to read `APP_VERSION` or other constants from `app_config.py`.

**Example:**
```powershell
# Read APP_VERSION from app_config.py
$versionMatch = Select-String -Path "app_config.py" -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
if (-not $versionMatch) {
    throw "APP_VERSION not found in app_config.py"
}
$AppVersion = $versionMatch.Matches.Groups[1].Value
Write-Host "[build] Version: $AppVersion"
```

[VERIFIED: PowerShell Select-String docs at learn.microsoft.com]

### Recommended Project Structure (Post Phase 1)

```
Virelo/
├── .gitignore              # NEW: Python + Node + build artifacts
├── CLAUDE.md               # REWRITTEN: build commands, conventions, footguns
├── LICENSE                  # NEW: MIT license
├── README.md               # NEW: what, why, build instructions
├── Virelo.spec             # RENAMED from "Windows Toolbox.spec"
├── app_config.py           # MODIFIED: add APP_VERSION, metadata constants
├── icon.ico
├── main.py                 # MODIFIED: remove deprecated Qt attrs, stale comments
├── requirements.txt        # UNCHANGED (Phase 3 adds pyproject.toml)
├── bridge.py               # UNCHANGED
├── webview.py              # UNCHANGED
├── settings.py             # UNCHANGED
├── settings_state.py       # UNCHANGED
├── snap_service.py         # UNCHANGED
├── capture_guard.py        # UNCHANGED
├── startup_shortcut.py     # UNCHANGED
├── theme.py                # MODIFIED: remove v2 comment
├── explorer_columns.py     # UNCHANGED
├── workers.py              # UNCHANGED
├── branding/               # UNCHANGED
├── installer/
│   └── virelo.iss          # MODIFIED: #ifndef for version, remove hardcoded version
├── frontend/
│   ├── package.json        # UNCHANGED (version is independent per D-01)
│   ├── vite.config.js      # MODIFIED: add define for __APP_VERSION__
│   └── src/
│       ├── app.jsx         # MODIFIED: use __APP_VERSION__ instead of hardcoded
│       ├── pages.jsx       # MODIFIED: use __APP_VERSION__ instead of hardcoded
│       ├── panels.jsx      # MODIFIED: use __APP_VERSION__ instead of hardcoded
│       ├── theme.jsx       # MODIFIED: remove v2 comment
│       ├── icons.jsx       # MODIFIED: remove v2 comment
│       └── primitives.jsx  # MODIFIED: remove v2 comment
└── scripts/
    ├── bootstrap.ps1       # NEW: create .venv, pip install, npm ci
    ├── clean.ps1           # NEW: remove build artifacts
    ├── build-frontend.ps1  # NEW: npm ci + npm run build + verify
    ├── build-app.ps1       # NEW: call build-frontend + PyInstaller + verify
    ├── build-installer.ps1 # REWRITTEN: call build-app + ISCC with version + verify
    └── verify-release.ps1  # NEW: check dist/ output integrity
```

### Anti-Patterns to Avoid

- **Importing app_config in the spec file:** The spec file runs in PyInstaller's build environment. Importing `app_config` may trigger imports of PySide6 or other packages that aren't available in the build context. Use regex to parse the version instead.
- **Hardcoding version in multiple places:** Every hardcoded version string is a future bug. The version must flow from `app_config.py` to everywhere else via build-time injection.
- **Modifying source files during builds:** The build scripts should NOT modify `vite.config.js`, `virelo.iss`, or any source file to inject the version. Use environment variables (Vite) and command-line flags (ISCC) instead.
- **Forgetting `$LASTEXITCODE` checks in PowerShell:** External commands (npm, python, pyinstaller, ISCC) do not trigger PowerShell's `$ErrorActionPreference = "Stop"`. Every external command call must be followed by an exit code check.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| .gitignore patterns | Custom ignore list from scratch | GitHub Python template + Node additions | GitHub's template covers edge cases (`.eggs/`, `*.egg-info/`, `*.manifest`, coverage artifacts) that are easy to forget |
| Version parsing in PowerShell | Custom string splitting | `Select-String` with regex capture group | Battle-tested, handles edge cases like whitespace variations |
| Build-time constants in frontend | Custom JS file generation | Vite `define` with `JSON.stringify` | Built into Vite, zero config, works in dev and prod |
| Installer version injection | Template file with sed/replace | ISCC `/D` flag with `#ifndef` fallback | Official Inno Setup mechanism, no temp files |
| MIT license text | Writing from memory | SPDX standard MIT text | Exact wording matters for legal validity |

**Key insight:** Every mechanism for version injection already exists as a first-class feature of its respective tool. No custom plumbing is needed.

## Common Pitfalls

### Pitfall 1: Vite `define` Without `JSON.stringify`

**What goes wrong:** Passing a raw string to Vite's `define` (e.g., `{ __APP_VERSION__: '1.5.0' }`) causes Vite to treat `1.5.0` as a JavaScript expression, not a string literal. The build succeeds but the frontend shows `undefined` or throws a ReferenceError.
**Why it happens:** Vite's `define` performs text replacement. Without `JSON.stringify`, the replacement value `1.5.0` is interpreted as the expression `1.5 - 0` which equals `1.5`, or as identifiers which are undefined.
**How to avoid:** Always use `JSON.stringify()`: `define: { __APP_VERSION__: JSON.stringify('1.5.0') }`.
**Warning signs:** Version displays as `NaN`, `undefined`, or a number instead of a semver string.

### Pitfall 2: PowerShell `$LASTEXITCODE` Not Checked After External Commands

**What goes wrong:** `npm run build` fails (e.g., syntax error in JSX) but the build script continues because `$ErrorActionPreference = "Stop"` only catches PowerShell cmdlet errors, not native command failures.
**Why it happens:** PowerShell distinguishes between cmdlet errors (thrown as exceptions) and native command failures (set `$LASTEXITCODE` but don't throw).
**How to avoid:** After every call to npm, python, pyinstaller, or ISCC, add: `if ($LASTEXITCODE -ne 0) { throw "command failed with exit code $LASTEXITCODE" }`.
**Warning signs:** Build scripts report success even when intermediate steps failed.

### Pitfall 3: ISCC `/D` Flag Ignored When `#define` Is Unconditional

**What goes wrong:** The .iss file has `#define MyAppVersion "1.4.2"` (unconditional). The build script passes `/DMyAppVersion=1.5.0`. The unconditional `#define` in the file overrides the command-line `/D`, so the installer is built with version "1.4.2".
**Why it happens:** The Inno Setup preprocessor processes directives in order. An unconditional `#define` after a command-line `/D` redefines the variable.
**How to avoid:** Use `#ifndef` guard: `#ifndef MyAppVersion` / `#define MyAppVersion "0.0.0-dev"` / `#endif`. The command-line `/D` sets the variable before the file is processed, so the `#ifndef` block is skipped.
**Warning signs:** Installer version doesn't match `app_config.py` version.

### Pitfall 4: PyInstaller Spec File Import Fails in Build Context

**What goes wrong:** The spec file does `from app_config import APP_VERSION`. This triggers the Python import system which may try to resolve other imports in `app_config.py` or its dependents. If `app_config.py` ever gains an import of PySide6 (even indirectly), the spec file fails with an ImportError.
**Why it happens:** Spec files run in PyInstaller's analysis phase, which has a restricted import environment.
**How to avoid:** Parse `app_config.py` with regex instead of importing it. The spec file should treat `app_config.py` as a text file, not a Python module.
**Warning signs:** `ImportError` or `ModuleNotFoundError` during `pyinstaller Virelo.spec`.

### Pitfall 5: Stale "Toolbox" References in Planning/Documentation Files

**What goes wrong:** After renaming everything, `grep -rn "Toolbox" .` still returns matches from `.planning/` docs, `CLAUDE.md`, or other documentation files. D-05 requires zero matches.
**Why it happens:** Planning documents and the auto-generated CLAUDE.md reference the old name. These are easy to overlook.
**How to avoid:** The grep verification in D-05 must exclude `.git/` but include `.planning/` and all markdown files. Update planning documents and CLAUDE.md as part of the cleanup.
**Warning signs:** D-05 verification grep returns matches after all source files are cleaned.

### Pitfall 6: Frontend `v2` Comment Removal Changes Module Semantics

**What goes wrong:** The first line of `theme.jsx`, `icons.jsx`, `panels.jsx`, `pages.jsx`, and `primitives.jsx` are comments like `// Pages for v2: Window Snap, Explorer, Shortcuts, General, About.` Removing them is safe BUT if a `"use strict"` directive or other pragma exists as the second line, removing the first-line comment could change its position semantics.
**Why it happens:** In JavaScript, directives like `"use strict"` are only treated as directives when they appear at the start of a script or function body.
**How to avoid:** When removing the `v2` comment line, verify the second line is not a directive. In this codebase, none of the files have `"use strict"` (React/Vite handles this), so removal is safe.
**Warning signs:** None expected -- this is a theoretical pitfall, not an active one in this codebase.

## Code Examples

### app_config.py with Version and Metadata Constants

```python
# Source: D-01, D-02 from CONTEXT.md
APP_NAME = "Virelo"
APP_DISPLAY_NAME = "Virelo"
APP_VERSION = "1.5.0"
APP_ID = "com.yusufqwareeq.virelo"
APP_EXECUTABLE_NAME = "Virelo.exe"
APP_DIST_DIR_NAME = "Virelo"
APP_PUBLISHER = "Yusuf Qwareeq"
APP_SUPPORT_URL = "https://github.com/yusufqwareeq/virelo"
APP_SETTINGS_ORG = "Yusuf Qwareeq"
APP_LOG_DIR = "Virelo"
APP_LOG_FILE = "virelo.log"
SETTINGS_GROUP = "Settings"

DEFAULTS = {
    # ... existing defaults unchanged ...
}
```

### virelo.iss with `#ifndef` Version Guard

```pascal
; Source: D-01 pattern, ISCC /D docs
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif
#define MyAppName "Virelo"
#define MyAppPublisher "Yusuf Qwareeq"
; ... rest of script uses {#MyAppVersion} as before ...
```

### vite.config.js with Version Define

```javascript
// Source: Vite define docs (vite.dev/config/shared-options.html#define)
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: './',
  define: {
    __APP_VERSION__: JSON.stringify(process.env.VITE_APP_VERSION || 'dev'),
  },
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

### Frontend Component Using Injected Version

```jsx
// Before (hardcoded):
<span>Virelo 1.4.2</span>

// After (injected at build time):
<span>Virelo {__APP_VERSION__}</span>
```

### PowerShell Version Extraction Pattern

```powershell
# Used by build-installer.ps1 and build-frontend.ps1
function Get-AppVersion {
    param([string]$ConfigPath = "app_config.py")
    $match = Select-String -Path $ConfigPath -Pattern 'APP_VERSION\s*=\s*"([^"]+)"'
    if (-not $match) {
        throw "APP_VERSION not found in $ConfigPath"
    }
    return $match.Matches.Groups[1].Value
}
```

### build-frontend.ps1: Setting VITE_APP_VERSION Environment Variable

```powershell
# Read version and pass to Vite via environment variable
$AppVersion = Get-AppVersion
$env:VITE_APP_VERSION = $AppVersion
Write-Host "[build-frontend] Version: $AppVersion"

Push-Location frontend
npm run build
if ($LASTEXITCODE -ne 0) { throw "npm run build failed" }
Pop-Location
```

## Stale Reference Inventory

This phase involves renaming. Here is the complete inventory of stale references found.

### "Windows Toolbox" References (IDENT-01, D-05)

| File | Line | Content | Action |
|------|------|---------|--------|
| `Windows Toolbox.spec` | filename | Spec file name | Rename to `Virelo.spec` |
| `scripts/build-installer.ps1` | 6 | `"Windows Toolbox.spec"` in Test-Path | Update to `Virelo.spec` |
| `scripts/build-installer.ps1` | 7 | `"Windows Toolbox.spec"` in error message | Update to `Virelo.spec` |
| `scripts/build-installer.ps1` | 11 | `"Windows Toolbox.spec"` in PyInstaller command | Update to `Virelo.spec` |
| `CLAUDE.md` | 40, 66 | References to `Windows Toolbox.spec` | CLAUDE.md will be rewritten entirely |
| `.planning/codebase/*` | multiple | References in planning docs | Update planning docs |

### Migration-Phase Comments (IDENT-04, D-04)

| File | Line | Content | Action |
|------|------|---------|--------|
| `main.py` | 908 | `Phase 7: UI replaced with React frontend...` | Remove or rewrite comment |
| `main.py` | 958 | `# Frameless + resizable window (SC-6).` | Remove `(SC-6)` reference |
| `main.py` | 991 | `# --- Bridge + WebView (Phase 7) ---` | Remove `(Phase 7)` reference |
| `main.py` | 1353 | `# Resizable window via WM_NCHITTEST (SC-6, IC-11)` | Remove `(SC-6, IC-11)` reference |
| `frontend/src/theme.jsx` | 1 | `// Theme tokens + context for Virelo v2.` | Remove `v2` label |
| `frontend/src/icons.jsx` | 1 | `// SVG icon component extracted from v2/primitives.jsx.` | Remove `v2` label |
| `frontend/src/panels.jsx` | 1 | `// Command palette (Ctrl/Cmd+K) for Virelo v2.` | Remove `v2` label |
| `frontend/src/pages.jsx` | 1 | `// Pages for v2: Window Snap, Explorer, Shortcuts, General, About.` | Remove `v2` label |
| `frontend/src/primitives.jsx` | 1 | `// Shared UI primitives for v2.` | Remove `v2` label |

### Hardcoded Version Strings (IDENT-03)

| File | Line | Content | Action |
|------|------|---------|--------|
| `installer/virelo.iss` | 5 | `#define MyAppVersion "1.4.2"` | Replace with `#ifndef`/`#endif` guard |
| `frontend/src/panels.jsx` | 127 | `<span>Virelo 1.4.2</span>` | Replace with `__APP_VERSION__` |
| `frontend/src/pages.jsx` | 232 | `{ v: '1.4.2', date: 'Apr 12', ... }` | Keep as changelog data (historical) |
| `frontend/src/pages.jsx` | 249 | `<span>Version 1.4.2</span>` | Replace with `__APP_VERSION__` |
| `frontend/src/app.jsx` | 105 | `v1.4.2 . up to date` | Replace with `__APP_VERSION__` |

**Note on pages.jsx:232:** The changelog array contains historical version entries. These should NOT be replaced with `__APP_VERSION__` -- they are historical data, not the current version display.

## .gitignore Pattern Reference

Complete pattern list for D-10, based on GitHub Python template with project-specific additions.

```gitignore
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
*.egg-info/
*.egg
dist/
build/
*.manifest

# Virtual environments
.venv/
venv/
ENV/

# Testing & coverage
.pytest_cache/
.coverage
htmlcov/
.ruff_cache/
.mypy_cache/

# Node / Frontend
frontend/node_modules/
frontend/dist/

# PyInstaller
*.spec.bak

# Inno Setup
installer/dist/

# Logs
*.log
crash.log

# IDE
.idea/
.vscode/
*.swp
*.swo

# OS
Thumbs.db
Desktop.ini
.DS_Store
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `Qt.AA_EnableHighDpiScaling` | Always enabled in Qt 6 | Qt 6.0 (2020) | Remove the `setAttribute` calls; they are no-ops |
| `Qt.AA_UseHighDpiPixmaps` | Always enabled in Qt 6 | Qt 6.0 (2020) | Remove the `setAttribute` calls; they are no-ops |
| Hardcoded version in .iss | ISCC `/D` flag injection | Always available in ISCC | Cleaner, no file modification needed |
| Single monolithic build script | Specialized scripts per stage | Pattern, not version-dependent | Better error isolation, reusable steps |

**Deprecated/outdated:**
- `Qt.AA_EnableHighDpiScaling`: Removed attribute -- always on in Qt 6. Will cause `AttributeError` in future PySide6 versions. [CITED: Qt 6 migration guide]
- `Qt.AA_UseHighDpiPixmaps`: Same as above.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | PowerShell `$ErrorActionPreference = "Stop"` does not catch native command failures | Pitfall 2, Pattern 3 | If it does catch them, the explicit `$LASTEXITCODE` checks are redundant but harmless |
| A2 | Importing `app_config` in the PyInstaller spec file may fail if it triggers PySide6 imports | Pitfall 4, Pattern 4 | If import works reliably, the regex approach is still safer and more portable |
| A3 | `bootstrap.ps1` creating a `.venv` is better than using system Python | Claude's Discretion area | Either approach works; .venv isolates dependencies and avoids polluting the system Python |

## Open Questions

1. **What version number to use after cleanup?**
   - What we know: Current version is 1.4.2 across all files
   - What's unclear: Should the version bump to 1.5.0 to signal the cleanup, or stay at 1.4.2 since no user-facing features changed?
   - Recommendation: Use 1.5.0 to signal that the build pipeline and identity are cleaned up. This is a meaningful change even though no features were added.

2. **Should `bootstrap.ps1` create a `.venv` or use system Python?**
   - What we know: The current `build-installer.ps1` uses `.venv\Scripts\python.exe`. No `.venv` currently exists. System Python is 3.13.
   - What's unclear: Whether the user prefers .venv isolation or system-level installation
   - Recommendation: Create `.venv` by default. It matches the existing build script's assumption and isolates project dependencies. The script should detect if `.venv` already exists and skip creation.

3. **Should `.planning/` docs be updated to remove "Toolbox" references?**
   - What we know: D-05 says grep must return zero matches for "Windows Toolbox" and "Toolbox"
   - What's unclear: Whether planning docs count as "source files" for D-05
   - Recommendation: Update `.planning/codebase/` docs that reference the old name, but the verification grep should exclude `.planning/` (these are planning artifacts, not source). CLAUDE.md IS source and must be cleaned.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Python | build-app.ps1, PyInstaller | Yes | 3.13.13 | -- |
| Node.js | build-frontend.ps1 | Yes | v24.15.0 | -- |
| npm | build-frontend.ps1 | Yes | 11.12.1 | -- |
| PowerShell | All build scripts | Yes | 5.1 (and 7.6) | -- |
| git | Version control | Yes | 2.54.0 | -- |
| PyInstaller | build-app.ps1 | No (not in global pip) | -- | `bootstrap.ps1` installs it into .venv |
| Inno Setup (ISCC.exe) | build-installer.ps1 | Not detected | -- | Script provides download URL on failure |
| PySide6 | Runtime dependency | Not in global pip | -- | `bootstrap.ps1` installs into .venv |

**Missing dependencies with no fallback:**
- None -- all missing dependencies can be installed via `bootstrap.ps1`

**Missing dependencies with fallback:**
- PyInstaller: Installed by `bootstrap.ps1` into `.venv` from `requirements.txt`
- Inno Setup: User must install manually; `build-installer.ps1` should provide download URL and fail early with clear message
- PySide6: Installed by `bootstrap.ps1` into `.venv` from `requirements.txt`

## Sources

### Primary (HIGH confidence)
- [Vite `define` documentation](https://vite.dev/config/shared-options.html#define) -- build-time constant injection syntax and JSON.stringify requirement
- [Inno Setup Preprocessor: Extended Command Line Compiler](https://jrsoftware.org/ishelp/topic_isppcc.htm) -- ISCC `/D` flag syntax
- [Inno Setup Preprocessor: #ifdef, #ifndef](https://jrsoftware.org/ishelp/topic_ifdef.htm) -- conditional define pattern
- [GitHub Python .gitignore template](https://github.com/github/gitignore/blob/main/Python.gitignore) -- standard Python ignore patterns
- [PyInstaller: Using Spec Files](https://pyinstaller.org/en/stable/spec-files.html) -- spec files are executable Python
- [PowerShell Select-String docs](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/select-string) -- regex capture group extraction
- Direct codebase analysis: `app_config.py`, `Windows Toolbox.spec`, `build-installer.ps1`, `virelo.iss`, `vite.config.js`, `main.py`, `webview.py`, `frontend/src/*.jsx` -- verified all stale references and version locations

### Secondary (MEDIUM confidence)
- `.planning/research/STACK.md` -- project-level stack research (verified tools and versions)
- `.planning/research/PITFALLS.md` -- project-level pitfalls (Pitfalls 1, 7, 13, 15 directly relevant)
- `.planning/research/FEATURES.md` -- feature dependency ordering

### Tertiary (LOW confidence)
- None -- all claims verified against official documentation or codebase inspection

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- no new libraries, all tools verified against official docs
- Architecture: HIGH -- version flow pattern uses first-class features of each tool (Vite define, ISCC /D, Python import/regex)
- Pitfalls: HIGH -- all pitfalls verified via codebase inspection and official documentation
- Build scripts: HIGH for pattern, MEDIUM for PowerShell edge cases (A1 assumption about $ErrorActionPreference)

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (stable domain -- build tools and patterns do not change rapidly)
