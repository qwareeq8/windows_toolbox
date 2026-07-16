# Building Virelo from Source

## Prerequisites

- Use Windows 10 version 1809 (build 17763) or later, or Windows 11, on x64.
- Install Git.
- Install 64-bit Python 3.12 to 3.14.
- Install Node.js 22 to 24.
- Install Inno Setup 6 when building the installer.

## Quick Start

```powershell
git clone https://github.com/qwareeq8/virelo.git
cd virelo
scripts/bootstrap.ps1 -Recreate
scripts/build-app.ps1
```

The one-folder application is created at `dist/Virelo/`. A release build requires a clean Git
working tree, including no untracked files. Use `-AllowDirty` only for a diagnostic build that
will not be released.

## Locked Environments

`requirements/build-constraints.txt` pins the Python application, test, and PyInstaller
environment. `frontend/package-lock.json` pins the Node.js dependency graph. Recreate the
environment after either file changes:

```powershell
scripts/bootstrap.ps1 -Recreate
```

The constraints file fixes version selection but does not authenticate package files. Review
dependency updates, use trusted indexes, retain build logs, and audit the resolved environment
for each release.

## Development Checks

Run the blocking source checks before building:

```powershell
.venv\Scripts\python.exe -m pip check
.venv\Scripts\python.exe -m ruff check .
.venv\Scripts\python.exe -m ruff format --check .
.venv\Scripts\python.exe -m mypy .
.venv\Scripts\python.exe -m pytest -q
cd frontend
npm audit --audit-level=high
npm run lint
npm run format:check
npm test
npm run build
cd ..
scripts/test-release-tools.ps1
```

## Build Pipeline

| Script | Purpose | Output |
|---|---|---|
| `scripts/bootstrap.ps1` | Create the constrained Python environment and run `npm ci`. | `.venv/` and `frontend/node_modules/` |
| `scripts/build-frontend.ps1` | Reinstall locked packages, lint, format-check, test, and build the versioned frontend. | `frontend/dist/` |
| `scripts/build-app.ps1` | Build the frontend, run PyInstaller, smoke-test the executable, and record provenance. | `dist/Virelo/` |
| `scripts/write-bundle-checksums.ps1` | Recompute the complete one-folder checksum inventory after an authorized signing step. | `dist/Virelo/bundle-files.sha256` |
| `scripts/build-installer.ps1` | Build the application and the Inno Setup installer. | `installer/dist/VireloSetup.exe` |
| `scripts/write-release-checksums.ps1` | Hash the final installer. The separate bundle inventory covers every portable-bundle file, including the executable and provenance record. | `installer/dist/CHECKSUMS.sha256` |
| `scripts/verify-release.ps1` | Validate static metadata, provenance, file inventories, checksums, packaged smoke tests, and Authenticode status. Add `-RequireSignature` to require valid signatures. | A pass, warning, or failure report. |
| `scripts/clean.ps1` | Remove generated application, installer, cache, coverage, and package metadata. | A clean generated-output state. |

`build-installer.ps1` runs `build-app.ps1` by default. After signing an existing application
bundle, use `build-installer.ps1 -SkipAppBuild` so the signed executable is not overwritten.

## Version Management

`APP_VERSION` in `virelo/app/config.py` is authoritative. The npm package and lock records are
compatibility-visible mirrors and must match it. Update them together:

```powershell
$Version = Read-Host "Release version"
cd frontend
npm version $Version --no-git-tag-version
cd ..
```

Then update `APP_VERSION` to the same value. The build derives the frontend constant, Python
package version, executable PE metadata, and installer metadata from those checked records.

## Development Mode

```powershell
$env:VIRELO_DEV = 1
cd frontend
npm run dev
```

In another terminal, run the application through the constrained environment:

```powershell
.venv\Scripts\python.exe main.py
```

## Verification Boundaries

Run the packaged smoke test directly when diagnosing a bundle:

```powershell
dist\Virelo\Virelo.exe --smoke-test
```

The smoke test checks startup dependencies, resources, QWebEngine construction, settings reads,
and bridge construction. It does not exercise installation, upgrade, rollback, uninstall,
window snapping, or live File Explorer registry changes. Those remain release acceptance tests
described in [RELEASE.md](RELEASE.md).
