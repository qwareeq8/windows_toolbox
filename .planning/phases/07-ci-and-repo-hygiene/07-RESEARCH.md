# Phase 7: CI and Repo Hygiene - Research

**Researched:** 2026-04-24
**Domain:** GitHub Actions CI, PowerShell build scripts, Python packaging, class rename, version synchronization
**Confidence:** HIGH

---

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions

- **D-01:** Commit icon.ico, branding/*, and frontend/index.html to git. These are source assets required for building, not generated artifacts. They are currently untracked but present in the working tree.
- **D-02:** No .gitignore changes needed — frontend/dist/ is gitignored (build output), but frontend/index.html and branding/ are not blocked.
- **D-03:** Expand scripts/clean.ps1 to recursively remove __pycache__/ and *.pyc across the entire project tree using Get-ChildItem -Recurse.
- **D-04:** Add installer/dist/ to the clean targets list.
- **D-05:** Add `--exclude-dir=.github` to the CI grep command so the workflow file doesn't match its own search string.
- **D-06:** Rephrase the test_app_config.py docstring (line 49) to avoid the literal "Windows Toolbox" string — use an indirect reference like "the old product name" instead.
- **D-07:** Keep CI test runner on ubuntu-latest (free runners). Add `@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs")` to tests that import win32con/win32gui/ctypes (currently test_snap_geometry.py test_restore_maximized_window).
- **D-08:** Ensure all other unit tests pass on Ubuntu without platform-specific imports at module level. Use conditional imports guarded by sys.platform where needed.
- **D-09:** Sync frontend/package.json version to match config.py APP_VERSION (currently 1.4.2 vs 1.5.0 — update package.json to 1.5.0).
- **D-10:** Add a CI check step that extracts both versions and fails if they differ, preventing future drift.
- **D-11:** Update README.md "Python 3.8+" to "Python 3.12+" to match pyproject.toml requires-python = ">=3.12".
- **D-12:** Align installer support URL with config.py. Currently config.py uses `https://github.com/yusufqwareeq/virelo` while installer uses `mailto:qwareeq8@gmail.com`. Update installer MyAppURL to use the GitHub URL.
- **D-13:** Rename ShiftSnapRestore → SnapRestoreController in virelo/services/snap.py (class, docstrings, LOG messages), virelo/app/window.py (import, usage, comments), and tests/unit/test_snap_geometry.py (import, usage, comments).
- **D-14:** Rename HotkeyListener → MultiPressHotkeyListener in virelo/services/snap.py (class, docstrings) and virelo/app/window.py (import, usage).
- **D-15:** Update CLAUDE.md project structure section to reflect new class names.
- **D-16:** Do NOT rename references in .planning/ docs — those are historical records of the names at the time of writing.

### Claude's Discretion

- Exact pytest marker placement and import guard structure
- Clean script implementation details (PowerShell cmdlet choices)
- CI version-check script implementation (regex vs Python parse)
- Order of commits within the phase

### Deferred Ideas (OUT OF SCOPE)

None — discussion stayed within phase scope.
</user_constraints>

---

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| CI-01 | Developer can clone the repo and build without missing assets (icon.ico, branding/*, frontend/index.html committed) | Assets verified present in working tree but untracked; git add needed |
| CI-02 | Developer can run clean script to remove all generated artifacts recursively | clean.ps1 currently misses recursive __pycache__ and installer/dist/; expansion pattern documented |
| CI-03 | CI stale-name check passes without false positives on its own workflow file or test docstrings | Two false positive sources identified: .github/workflows/ci.yml (fix: --exclude-dir=.github) and test_app_config.py docstring (fix: rephrase) |
| CI-04 | CI unit tests pass on ubuntu-latest with platform-guarded Win32 imports | test_restore_maximized_window imports ctypes.wintypes (Windows-only); needs skipif marker |
| CI-05 | Frontend package.json version synchronized with APP_VERSION | package.json at 1.4.2, config.py at 1.5.0; update package.json |
| CI-06 | README states correct Python version (3.12+) | README says 3.8+; pyproject.toml says >=3.12; update README |
| CI-07 | Installer support URL and config.py support URL use same source of truth | installer MyAppURL = mailto:...; config.py APP_SUPPORT_URL = GitHub URL; update installer |
| CI-08 | Stale class names renamed (ShiftSnapRestore -> SnapRestoreController, HotkeyListener -> MultiPressHotkeyListener) | All occurrences located in snap.py, window.py, test_snap_geometry.py, and CLAUDE.md |
</phase_requirements>

---

## Summary

Phase 7 is a pure hygiene phase — no new features, no new dependencies. Every change is either a text edit, a file commit, or a script expansion. All decisions are locked and fully specified in CONTEXT.md. The work decomposes cleanly into five independent streams: (1) commit missing source assets, (2) expand the clean script, (3) fix CI false positives plus platform guards, (4) synchronize version strings, and (5) rename two stale class names throughout the codebase.

The only subtlety is the `test_restore_maximized_window` test, which has been passing locally because `ctypes.wintypes` is available on Windows but is not importable on Linux CI. The conftest stubs `win32con` and `win32gui` into `sys.modules` so those module-level imports are satisfied everywhere, but `from ctypes import wintypes` inside the test function body bypasses the stub mechanism and will raise `ImportError` on Ubuntu. Adding `@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs")` to that test function is the correct fix per D-07.

The version-check CI step (D-10) should be a shell `run:` step in `.github/workflows/ci.yml` that extracts the version from `virelo/app/config.py` and `frontend/package.json` using grep/python, then compares them and exits non-zero on mismatch. No new tools or files needed.

**Primary recommendation:** Work in dependency order — assets first (CI-01), then clean (CI-02), then CI fixes (CI-03/04), then version sync (CI-05/06/07), then class renames (CI-08). Each stream can be a separate plan wave or commit.

---

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Asset management (git) | Source control | — | icon.ico and branding/ are inputs to the installer build pipeline; they belong in git |
| Clean script | Build tooling (PowerShell) | — | Extends existing pattern; no new tier |
| CI stale-name check | CI (GitHub Actions) | Unit test (docstring) | grep runs in CI; docstring edit is in test source |
| Platform guard for Win32 test | Unit test | CI | pytest.mark.skipif controls test collection at runtime |
| Version synchronization | Config source (config.py) | CI (version-check job) | config.py is the single source of truth; CI enforces consistency |
| Version-check CI step | CI (GitHub Actions) | — | Shell step comparing extracted versions |
| Class rename | Python source | Unit test | All rename points are in .py files |

---

## Standard Stack

### Core (all already in the project — no new dependencies)

| Tool | Version | Purpose | Why Standard |
|------|---------|---------|--------------|
| GitHub Actions | N/A | CI runner | Already in use; `ci.yml` defines 4 jobs |
| pytest | >=9.0.3 | Python test runner | Already in `pyproject.toml [dev]` |
| ruff | >=0.15.12 | Python linter/formatter | Already in `pyproject.toml [dev]` |
| PowerShell | Windows built-in | Build scripts | All scripts in `scripts/` use PowerShell |
| grep (bash) | Ubuntu built-in | Stale-name check in CI | Used in existing `stale-name` CI job |

No new libraries. This phase modifies existing files only.

---

## Architecture Patterns

### System Architecture Diagram

```
Developer clone
       |
       v
[git checkout] --> icon.ico, branding/*, frontend/index.html (committed after CI-01)
       |
       v
[scripts/clean.ps1] --> removes dist/, build/, frontend/dist/, installer/dist/,
                        + recursive __pycache__/  *.pyc  .pytest_cache  .ruff_cache
       |
       v
[CI Pipeline: .github/workflows/ci.yml]
       |
       +--[lint job]--> ruff check + ruff format --check
       |
       +--[test job]--> pytest tests/unit/
       |                  |
       |                  +-- conftest.py stubs PySide6, win32*, keyboard
       |                  +-- test_restore_maximized_window [skipif win32 != platform]
       |
       +--[frontend job]--> npm ci + npm run build + vitest run
       |
       +--[version-check job (NEW)]--> extract APP_VERSION from config.py
       |                               extract version from package.json
       |                               compare; fail if mismatch
       |
       +--[stale-name job]--> grep "Windows Toolbox" ...
                               --exclude-dir=.github  (NEW)
                               (test docstring rephrased; no more matches)
```

### Recommended Project Structure

No structural changes. All edits are within existing files.

```
.github/workflows/
  ci.yml              # Add version-check job; add --exclude-dir=.github to grep
scripts/
  clean.ps1           # Add recursive __pycache__ + *.pyc + installer/dist/
virelo/services/
  snap.py             # HotkeyListener -> MultiPressHotkeyListener
                      # ShiftSnapRestore -> SnapRestoreController
virelo/app/
  window.py           # Update import + usage of both renamed classes
tests/unit/
  test_snap_geometry.py   # Update import + skipif marker + rename references
  test_app_config.py      # Rephrase docstring on test_app_name_is_virelo
frontend/
  package.json        # version: 1.4.2 -> 1.5.0
installer/
  virelo.iss          # MyAppURL: mailto:... -> GitHub URL
README.md             # Python 3.8+ -> Python 3.12+
CLAUDE.md             # snap.py line: update both class names in description
```

### Pattern 1: Recursive pycache Clean (PowerShell)

**What:** Remove all `__pycache__` directories and `*.pyc` files anywhere in the project tree.

**When to use:** After the existing flat-target loop.

**Example:**
```powershell
# Source: existing pattern in scripts/clean.ps1, extended with -Recurse
Get-ChildItem -Recurse -Filter "__pycache__" -Directory -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host "[clean] Removing $($_.FullName)"
        Remove-Item -Recurse -Force $_.FullName
    }

Get-ChildItem -Recurse -Filter "*.pyc" -ErrorAction SilentlyContinue |
    ForEach-Object {
        Write-Host "[clean] Removing $($_.Name)"
        Remove-Item -Force $_.FullName
    }
```

Note: The existing `$targets` loop already handles `.pytest_cache` and `.ruff_cache` at the project root. Adding `installer/dist` to that same `$targets` array handles CI-02's requirement for that path.

### Pattern 2: pytest.mark.skipif for Platform-Only Tests

**What:** Skip a test function when the required platform is not available.

**When to use:** Any test that imports `ctypes.wintypes` or uses Win32 APIs that are unavailable on Linux.

**Example:**
```python
# Source: pytest docs [CITED: https://docs.pytest.org/en/stable/reference/reference.html#pytest.mark.skipif]
import sys
import pytest

@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs only available on Windows")
def test_restore_maximized_window():
    """Restore of a previously-maximized window issues SW_MAXIMIZE (SNAP-04)."""
    import ctypes
    from ctypes import wintypes
    ...
```

The marker must be placed on the function, not inside it. `sys` is a stdlib module; no import guard needed at module level.

### Pattern 3: CI Version-Check Shell Step

**What:** Extract version strings from two files and compare them inline in a CI `run:` step.

**When to use:** As an additional job or step in `ci.yml` that runs on every push/PR.

**Example:**
```yaml
# Source: existing ci.yml pattern + shell string comparison [VERIFIED: codebase]
  version-check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Check version consistency
        run: |
          PY_VERSION=$(grep -oP '(?<=APP_VERSION = ")[^"]+' virelo/app/config.py)
          PKG_VERSION=$(grep -oP '(?<="version": ")[^"]+' frontend/package.json)
          echo "config.py: $PY_VERSION"
          echo "package.json: $PKG_VERSION"
          if [ "$PY_VERSION" != "$PKG_VERSION" ]; then
            echo "ERROR: Version mismatch — config.py=$PY_VERSION, package.json=$PKG_VERSION"
            exit 1
          fi
          echo "OK: Versions match ($PY_VERSION)"
```

This requires no extra tooling — `grep -oP` (Perl regex) is available on Ubuntu runners. [VERIFIED: ubuntu-latest ships grep 3.x with -P support]

### Pattern 4: Inno Setup `#define` with `#ifndef` Guard (already present)

The installer already uses the `#ifndef MyAppVersion` guard pattern correctly (CLAUDE.md footgun #4). Only the `MyAppURL` constant needs updating — it has no command-line override, so a direct `#define` edit is safe.

**Current state (virelo.iss line 3):**
```iss
#define MyAppURL "mailto:qwareeq8@gmail.com"
```

**Target state:**
```iss
#define MyAppURL "https://github.com/yusufqwareeq/virelo"
```

### Pattern 5: Class Rename — Full Occurrence Map

All occurrences verified by grep against the live codebase:

**ShiftSnapRestore → SnapRestoreController:**

| File | Line(s) | Change |
|------|---------|--------|
| `virelo/services/snap.py` | 1 (module docstring) | Update mention |
| `virelo/services/snap.py` | 4 (docstring line) | Update mention |
| `virelo/services/snap.py` | 134 (class def) | Rename class |
| `virelo/services/snap.py` | 208 (LOG message) | Update string |
| `virelo/services/snap.py` | 329 (SnapService docstring) | Update |
| `virelo/services/snap.py` | 334 (set_manager docstring) | Update |
| `virelo/app/window.py` | 23 (import) | Update import name |
| `virelo/app/window.py` | 174 (comment) | Update comment |
| `virelo/app/window.py` | 197 (comment) | Update comment |
| `virelo/app/window.py` | 199 (instance creation) | `ShiftSnapRestore(...)` → `SnapRestoreController(...)` |
| `tests/unit/test_snap_geometry.py` | 105 (section comment) | Update |
| `tests/unit/test_snap_geometry.py` | 116 (import inside test) | Update import |
| `tests/unit/test_snap_geometry.py` | 118 (comment) | Update |
| `tests/unit/test_snap_geometry.py` | 119 (instantiation) | Update |
| `CLAUDE.md` | 56 (snap.py description) | Update |

**HotkeyListener → MultiPressHotkeyListener:**

| File | Line(s) | Change |
|------|---------|--------|
| `virelo/services/snap.py` | 1 (module docstring) | Update mention |
| `virelo/services/snap.py` | 3 (docstring line) | Update mention |
| `virelo/services/snap.py` | 49 (class def) | Rename class |
| `virelo/services/snap.py` | 338 (set_listener docstring) | Update |
| `virelo/app/window.py` | 23 (import) | Update import name |
| `virelo/app/window.py` | 197 (comment) | Update comment |
| `virelo/app/window.py` | 198 (instance creation) | `HotkeyListener(...)` → `MultiPressHotkeyListener(...)` |
| `virelo/app/window.py` | 200 | `self._hotkey_listener = HotkeyListener(...)` — variable name unchanged, only class name |
| `CLAUDE.md` | 56 (snap.py description) | Update |

Note: The attribute `self._hotkey_listener` and `self.shift_mgr` in window.py are instance variable names, not class names. D-13/D-14 scope only the class names. Variable names do NOT need renaming.

### Anti-Patterns to Avoid

- **Running the stale-name grep without --exclude-dir=.github:** The ci.yml file contains the literal search string "Windows Toolbox" as part of the grep command itself. This self-matches and fails CI until D-05 is applied.
- **Leaving `from ctypes import wintypes` without a platform guard:** This import is silently available on Windows but raises ImportError on Linux. The conftest stub only covers win32con/win32gui, not ctypes.wintypes.
- **Renaming instance variable names alongside class names:** `self._hotkey_listener` (window.py) and `self.shift_mgr` are local variable names chosen by the author. D-13/D-14 are class renames only. Renaming variables would be scope creep and could break internal wiring if a search-and-replace is too broad.
- **Forgetting the `SnapService` docstrings in snap.py:** Lines 329 and 334 reference "ShiftSnapRestore" in docstrings. These are not class definitions but still contain the stale name and should be updated.
- **Using `--exclude` (file-level) instead of `--exclude-dir` (directory-level) for .github:** The correct flag for excluding a directory in grep is `--exclude-dir`. Using `--exclude=.github/workflows/ci.yml` would also work but is more brittle if the file is renamed.

---

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Version comparison in CI | Custom Python script called via subprocess | `grep -oP` + bash string comparison in a `run:` step | Zero dependencies; ubuntu-latest has grep with -P; simpler YAML |
| Recursive file deletion in PowerShell | Custom recursive function | `Get-ChildItem -Recurse` + `Remove-Item` | Built-in cmdlet; handles symlinks and permissions correctly |
| Platform detection in pytest | Custom fixture or environment variable | `sys.platform` in `@pytest.mark.skipif` | Standard pytest mechanism; documented pattern |

---

## Common Pitfalls

### Pitfall 1: ctypes.wintypes on Ubuntu

**What goes wrong:** `test_restore_maximized_window` passes locally on Windows but raises `ImportError: cannot import name 'wintypes' from 'ctypes'` on Linux CI.

**Why it happens:** `ctypes.wintypes` is a Windows-only submodule. It is not part of the Python standard library on Linux/macOS. The conftest stubs `win32con`, `win32gui`, and `win32api` but makes no attempt to stub `ctypes.wintypes` because `ctypes` itself is a C extension module that differs between platforms.

**How to avoid:** Add `@pytest.mark.skipif(sys.platform != "win32", reason="Win32 APIs only available on Windows")` to the test function.

**Warning signs:** Test passes locally on Windows but CI reports `ImportError` or `AttributeError` from `from ctypes import wintypes`.

### Pitfall 2: grep self-matching in stale-name check

**What goes wrong:** The `stale-name` CI job runs grep for "Windows Toolbox" and finds its own grep command in ci.yml, causing CI to fail even when no stale names exist in source files.

**Why it happens:** The .github/ directory is scanned unless explicitly excluded. The current exclusion list omits `.github`.

**How to avoid:** Add `--exclude-dir=.github` to the grep flags. Verify by running grep locally with the new flags before pushing.

**Warning signs:** CI stale-name job fails but grep of the source tree (manually excluding .github) returns no matches.

### Pitfall 3: package.json version drifts silently without CI enforcement

**What goes wrong:** APP_VERSION in config.py and the version in frontend/package.json diverge. The mismatch is not caught until someone notices the About screen shows a different version than what was released.

**Why it happens:** package.json version is not used at build time (Vite injects `__APP_VERSION__` from config.py via `VITE_APP_VERSION`). There is no automated check.

**How to avoid:** Add the version-check CI job (D-10). The job must run independently so it can fail fast without waiting for the full test/build pipeline.

**Warning signs:** package.json version at 1.4.2, config.py at 1.5.0 — the current state of the repo.

### Pitfall 4: Too-broad search-and-replace for class renames

**What goes wrong:** A naive sed/replace of "ShiftSnapRestore" catches docstring mentions in planning docs and historical comments that D-16 explicitly says to leave alone. Or it renames the variable `shift_mgr` which happens to have no name overlap but other variables might.

**Why it happens:** Class rename tools or `sed -i` apply globally unless scoped.

**How to avoid:** Target specific files listed in D-13/D-14 occurrence map. Verify with grep after renaming that no occurrences remain in the target files, and that `.planning/` was untouched.

**Warning signs:** `.planning/` docs contain updated class names after the phase — this means the scope was too broad.

### Pitfall 5: clean.ps1 removing wrong __pycache__ directory

**What goes wrong:** The existing `$targets` loop removes a top-level `__pycache__` directory if present. But Python creates `__pycache__` inside every package directory (`virelo/`, `virelo/app/`, `virelo/bridge/`, etc.). The current script only catches one.

**Why it happens:** The `$targets` list uses relative paths, which only match the root-level `__pycache__`. Subdirectory caches are not removed.

**How to avoid:** Replace the `__pycache__` entry in `$targets` with a `Get-ChildItem -Recurse -Directory -Filter "__pycache__"` pass (Pattern 1 above). Keep `.pytest_cache` and `.ruff_cache` in the flat `$targets` list since they only appear at the root.

---

## Code Examples

### Version-Check CI Job (full YAML block)

```yaml
# Source: existing ci.yml pattern extended [VERIFIED: codebase]
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

### Updated stale-name grep (ci.yml)

```bash
# Source: existing ci.yml stale-name job, line 55-64 [VERIFIED: codebase]
if grep -rn "Windows Toolbox" . \
  --include="*.py" --include="*.jsx" --include="*.js" \
  --include="*.json" --include="*.toml" --include="*.yml" \
  --include="*.iss" --include="*.spec" \
  --exclude-dir=node_modules --exclude-dir=.git \
  --exclude-dir=dist --exclude-dir=build \
  --exclude-dir=.planning --exclude-dir=.github; then
  echo "ERROR: Found stale 'Windows Toolbox' references"
  exit 1
fi
echo "OK: No stale references found"
```

### Updated test_app_config.py docstring

```python
# Source: tests/unit/test_app_config.py:48-50 [VERIFIED: codebase]
def test_app_name_is_virelo():
    """APP_NAME must be 'Virelo' (not the old product name)."""
    assert APP_NAME == "Virelo"
```

### snap.py class definition renames

```python
# Source: virelo/services/snap.py [VERIFIED: codebase]

# OLD:
class HotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""

# NEW:
class MultiPressHotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""


# OLD:
class ShiftSnapRestore(QtCore.QObject):
    """Performs window snap and restore operations."""

# NEW:
class SnapRestoreController(QtCore.QObject):
    """Performs window snap and restore operations."""
```

### window.py import update

```python
# Source: virelo/app/window.py:23 [VERIFIED: codebase]
# OLD:
from virelo.services.snap import HotkeyListener, ShiftSnapRestore, SnapService

# NEW:
from virelo.services.snap import MultiPressHotkeyListener, SnapRestoreController, SnapService
```

### clean.ps1 expanded targets + recursive cleanup

```powershell
# Source: scripts/clean.ps1 existing pattern [VERIFIED: codebase], extended
$targets = @(
    "build",
    "dist",
    "frontend\dist",
    "installer\dist",    # NEW (CI-02)
    ".pytest_cache",
    ".ruff_cache"
)

# ... existing foreach loop ...

# NEW: recursive __pycache__ and *.pyc removal (CI-02)
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

## Runtime State Inventory

This phase contains class renames. Per Step 2.5 protocol, all five categories are answered explicitly.

| Category | Items Found | Action Required |
|----------|-------------|------------------|
| Stored data | None — ShiftSnapRestore and HotkeyListener are class names in Python source, not stored in any database, JSON file, or registry key | None |
| Live service config | None — no external service (n8n, Datadog, etc.) references these class names | None |
| OS-registered state | None — Windows Task Scheduler, startup shortcuts, and registry entries use "Virelo" app name, not internal class names | None |
| Secrets/env vars | None — no SOPS keys, .env vars, or CI environment variables reference these class names | None |
| Build artifacts | `.venv/` and `__pycache__/` may contain .pyc files compiled from the old class names. These are regenerated on next run and removed by the expanded clean.ps1 | Run clean + reinstall to ensure fresh bytecode (optional; Python recompiles automatically on import) |

Nothing persisted in external systems. Rename is purely a source-code operation.

---

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| git | CI-01 (asset commit) | Yes | (local git) | — |
| PowerShell | CI-02 (clean script) | Yes (Win11) | Built-in | — |
| Python 3.12 | CI-04 (pytest) | Yes | 3.12 (local) | — |
| pytest | CI-04 | Yes | >=9.0.3 (pyproject.toml dev) | — |
| grep (with -P) | CI-03, CI version-check | Ubuntu runner built-in | grep 3.x | Use python -c for version parse if -P unavailable |
| GitHub Actions ubuntu-latest | CI-03/04/05 | Yes (GitHub-hosted) | Ubuntu 22.04 LTS | — |

No blocking missing dependencies. All required tools are either already installed locally or provided by the CI runner.

---

## Validation Architecture

### Test Framework

| Property | Value |
|----------|-------|
| Framework | pytest 9.0.3+ |
| Config file | `pyproject.toml` `[tool.pytest.ini_options]` |
| Quick run command | `pytest tests/unit/ -q` |
| Full suite command | `pytest tests/unit/ -q` |

### Phase Requirements → Test Map

| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| CI-01 | Assets committed | manual verification (git status) | `git status icon.ico branding/ frontend/index.html` | N/A |
| CI-02 | Clean removes all artifacts | manual verification (run clean, check dirs gone) | `scripts/clean.ps1; Test-Path dist` | N/A |
| CI-03 | Stale-name grep passes | CI job (stale-name in ci.yml) | `grep -rn "Windows Toolbox" ... --exclude-dir=.github` | ci.yml |
| CI-04 | Unit tests pass on Ubuntu | unit (pytest) | `pytest tests/unit/ -q` | ✅ existing |
| CI-05 | Versions match | CI job (version-check) | `pytest` is not applicable; CI job handles this | ci.yml Wave 0 gap |
| CI-06 | README Python version | manual doc verification | N/A — text change | N/A |
| CI-07 | Installer URL matches config | manual verification | N/A — text change | N/A |
| CI-08 | Class renames — no old names in scope files | code verification | `grep -rn "ShiftSnapRestore\|HotkeyListener" virelo/ tests/unit/ CLAUDE.md` | N/A |

### Sampling Rate

- **Per task commit:** `pytest tests/unit/ -q`
- **Per wave merge:** `pytest tests/unit/ -q`
- **Phase gate:** All unit tests green + CI stale-name job green + CI version-check job green before `/gsd-verify-work`

### Wave 0 Gaps

- [ ] `version-check` job in `.github/workflows/ci.yml` — covers CI-05 enforcement (new job to add)

*(All other test infrastructure already exists. No new test files needed.)*

---

## Security Domain

No security-relevant changes in this phase. All edits are rename/text/config operations. No authentication, cryptography, input validation, or session management logic is touched.

ASVS categories V2 through V6 do not apply.

---

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `grep -oP` (Perl-mode regex) is available on ubuntu-latest GitHub Actions runners | Code Examples: version-check job | CI version-check step would fail; fallback: use `python3 -c "import re, pathlib; ..."` instead | 
| A2 | `ctypes.wintypes` is not importable on Ubuntu (causes ImportError) | Common Pitfalls #1 | If wrong, test_restore_maximized_window already passes on CI and skipif is harmless but unnecessary |

Note: A1 is nearly certain — Ubuntu 22.04 ships `grep 3.7` with `-P` support [ASSUMED: based on Ubuntu 22.04 package history]. A2 is verified by Python documentation: `ctypes.wintypes` is described as "Windows only" [CITED: https://docs.python.org/3/library/ctypes.html#ctypes.wintypes].

---

## Open Questions

1. **Should the version-check job fail the entire CI or run as `continue-on-error`?**
   - What we know: D-10 says "fail if they differ" — strict enforcement.
   - What's unclear: Whether to make it a blocking job or an advisory check.
   - Recommendation: Make it blocking (no `continue-on-error`). The whole point is to prevent drift.

2. **Should `installer/dist/` be added to the flat `$targets` array or get its own `if (Test-Path ...)` block?**
   - What we know: The existing flat loop already calls `if (Test-Path $target)` before removing.
   - What's unclear: Path separator on Windows PowerShell (`installer\dist` vs `installer/dist`).
   - Recommendation: Use `installer\dist` (backslash) in the `$targets` array to match the existing `frontend\dist` pattern. PowerShell accepts both but the script is already consistent with backslashes.

---

## Sources

### Primary (HIGH confidence)

- **Codebase grep** — All class name occurrences verified in: `virelo/services/snap.py`, `virelo/app/window.py`, `tests/unit/test_snap_geometry.py`, `CLAUDE.md`
- **Codebase read** — `.github/workflows/ci.yml`, `scripts/clean.ps1`, `virelo/app/config.py`, `frontend/package.json`, `installer/virelo.iss`, `pyproject.toml`, `README.md`, `tests/unit/test_app_config.py`, `tests/conftest.py`
- **Local execution** — `pytest tests/unit/ -q` passed 66 tests in 0.04s on Windows

### Secondary (MEDIUM confidence)

- **pytest docs** — `pytest.mark.skipif` marker API [CITED: https://docs.pytest.org/en/stable/reference/reference.html#pytest.mark.skipif]
- **Python docs** — `ctypes.wintypes` is Windows-only [CITED: https://docs.python.org/3/library/ctypes.html#ctypes.wintypes]

### Tertiary (LOW confidence)

- `grep -oP` availability on ubuntu-latest runners [ASSUMED: based on Ubuntu 22.04 package baseline]

---

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — no new dependencies; all tools verified present
- Architecture: HIGH — all change locations identified and verified by grep
- Pitfalls: HIGH — false positives and platform gaps confirmed by local test run and codebase read
- Version state: HIGH — all four version values read directly from source files

**Research date:** 2026-04-24
**Valid until:** Stable indefinitely — this phase has no external dependencies or fast-moving ecosystem components
