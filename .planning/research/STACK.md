# Technology Stack: Hardening & Professionalization Tooling

**Project:** Virelo (PySide6 + React + QWebEngineView hybrid desktop app)
**Researched:** 2026-04-24
**Scope:** Testing, linting, CI/CD, build pipeline, WebEngine hardening -- NOT the existing app stack

## Executive Summary

The existing Virelo app stack (PySide6, React 19, Vite, QWebChannel, PyInstaller, Inno Setup) is established and correct. This document covers the tooling layer that needs to be added on top: quality gates, testing frameworks, build automation, and hardening patterns. Every tool below was selected for this specific context -- a personal-use Windows desktop utility with a Python backend and a thin React settings UI.

---

## Project Configuration

### Migrate to pyproject.toml

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| pyproject.toml (setuptools) | PEP 621 | Single source of truth for project metadata, dependencies, tool config | Replaces `requirements.txt` + scattered tool configs. Standard since 2022, universally supported. Every tool below reads config from pyproject.toml. |

**Confidence:** HIGH -- PEP 621 is the Python packaging standard; all tools below support it.

The current `requirements.txt` should be replaced with a `pyproject.toml` that holds:
- Project metadata and version (single source of truth)
- Runtime dependencies (what's in requirements.txt now)
- Optional dependency groups: `[project.optional-dependencies]` for `dev`, `test`, `lint`
- Tool configuration sections for ruff, mypy, pytest

Build dependencies (PyInstaller, pyinstaller-hooks-contrib) move to `[project.optional-dependencies] build = [...]` since they are not runtime requirements.

### Pin the Python Version

| Technology | Value | Purpose | Why |
|------------|-------|---------|-----|
| `.python-version` | `3.12` | Pin development Python version | PySide6 6.8+ supports 3.10-3.14. Pin 3.12 for stability -- it is the most widely tested with PySide6 and PyInstaller. |

**Confidence:** MEDIUM -- 3.12 is the safe conservative choice; 3.13 works too but 3.12 has longer track record with the dependency set.

---

## Python Linting & Formatting

### Ruff (Linter + Formatter)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| ruff | >=0.15.11 | Python linting AND formatting | Replaces flake8, isort, black, pyupgrade, bandit in one tool. Written in Rust, sub-second on any codebase. Single config in pyproject.toml. |

**Confidence:** HIGH -- Ruff is the dominant Python linting tool as of 2026. Actively maintained by Astral (same team as uv). v0.15.11 released 2026-04-16.

**Configuration (pyproject.toml):**
```toml
[tool.ruff]
target-version = "py312"
line-length = 100

[tool.ruff.lint]
select = [
    "E",     # pycodestyle errors
    "W",     # pycodestyle warnings
    "F",     # pyflakes
    "I",     # isort
    "N",     # pep8-naming
    "UP",    # pyupgrade
    "B",     # flake8-bugbear
    "SIM",   # flake8-simplify
    "TCH",   # flake8-type-checking
    "RUF",   # ruff-specific rules
]
ignore = ["E501"]  # line length handled by formatter

[tool.ruff.format]
quote-style = "double"
```

**What NOT to use:**
- **flake8** -- Ruff implements all flake8 rules 10-100x faster. No reason to run both.
- **black** -- `ruff format` is a drop-in replacement with identical output, zero config.
- **isort** -- Ruff's `I` rules handle import sorting.
- **pylint** -- Slow, complex config, overlaps heavily with Ruff. Not worth the overhead for a personal project.

### mypy (Type Checker)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| mypy | >=1.20.2 | Static type checking | Catches type errors before runtime. Critical for Win32/COM code where wrong types cause silent failures. |

**Confidence:** HIGH -- mypy 1.20.2 released 2026-04-21. Standard Python type checker, well-maintained.

**Configuration (pyproject.toml):**
```toml
[tool.mypy]
python_version = "3.12"
warn_return_any = true
warn_unused_configs = true
disallow_untyped_defs = false   # start permissive, tighten over time
check_untyped_defs = true
ignore_missing_imports = true   # pywin32/comtypes stubs are incomplete

[[tool.mypy.overrides]]
module = ["win32gui", "win32api", "win32con", "win32event", "win32com.*", "pythoncom", "pywintypes", "comtypes.*", "keyboard"]
ignore_missing_imports = true
```

**Why `disallow_untyped_defs = false`:** The codebase has zero type annotations today. Starting strict would produce hundreds of errors and block progress. Enable gradually.

**What NOT to use:**
- **pyright** -- Excellent tool, but mypy has better ecosystem support for pre-commit hooks and CI. For a project adding types incrementally, mypy's gradual typing story is stronger.
- **ty (Astral)** -- Too new (early 2026 preview). Not production-ready yet for a project that needs stability.

---

## Python Testing

### pytest (Test Runner)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| pytest | >=9.0.2 | Python test runner | De facto standard. Plugin ecosystem covers Qt, coverage, fixtures. v9.0.2 released 2026-04-07. |
| pytest-cov | >=7.1.0 | Coverage reporting | Integrates coverage.py with pytest. v7.1.0 released 2026-03-21. |
| pytest-qt | >=4.5.0 | PySide6/Qt widget testing | Provides `qtbot` fixture for simulating Qt widget interaction. Headless-capable for CI. |

**Confidence:** HIGH -- pytest is the unquestioned standard. pytest-qt is the only serious option for testing PySide6 code.

**Configuration (pyproject.toml):**
```toml
[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "--cov=virelo --cov-report=term-missing --cov-report=html -q"
qt_api = "pyside6"
```

**Test organization:**
```
tests/
  test_settings.py       # SettingsState validation, defaults, coercion
  test_snap_service.py   # Snap logic with mocked Win32 calls
  test_bridge.py         # Bridge method dispatch (mock QWebChannel)
  test_app_config.py     # Config defaults, version string
  test_theme.py          # Theme resolution logic
```

**What to test first:** Pure logic -- `settings_state.py` (validation, coercion), `app_config.py` (defaults), `theme.py` (theme resolution). These have no Win32 dependencies and can run anywhere.

**What to mock:** All Win32 API calls (`win32gui`, `ctypes.windll`), COM objects, `keyboard` module. These are integration boundaries, not unit test targets.

**What NOT to use:**
- **unittest** -- pytest runs unittest tests anyway, but pytest's fixture model is cleaner. No reason to write new tests with unittest.
- **tox** -- Overkill for a single-platform, single-Python-version project. Run pytest directly in CI.

---

## Frontend Linting & Testing

### Biome (Frontend Linter + Formatter)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| biome | >=2.3 (dev) | JavaScript/JSX linting and formatting | Single tool replacing ESLint + Prettier. Written in Rust, instant. Zero config needed for JSX. |

**Confidence:** MEDIUM-HIGH -- Biome v2.3 is mature and stable for JSX projects. Chose Biome over ESLint because of the ESLint 10 transition mess.

**Why NOT ESLint:**
- ESLint 10.0.0 (released 2026-02) removed legacy config and broke `eslint-plugin-react` compatibility (GitHub issue #3977 still open as of April 2026).
- `eslint-plugin-react-hooks` also lacks ESLint 10 peer dependency support (facebook/react#35758).
- For a small settings UI with ~8 JSX files, the ESLint plugin ecosystem adds no value. Biome covers JSX/React rules natively.

**Configuration (`biome.json` in `frontend/`):**
```json
{
  "$schema": "https://biomejs.dev/schemas/2.3/schema.json",
  "organizeImports": { "enabled": true },
  "linter": { "enabled": true },
  "formatter": {
    "enabled": true,
    "indentStyle": "space",
    "indentWidth": 2
  }
}
```

### Vitest (Frontend Test Runner)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| vitest | >=4.1.4 (dev) | JavaScript test runner | Built on Vite -- shares the same config, same transforms. Zero additional bundler config needed. |
| @testing-library/react | >=16.3.2 (dev) | React component testing | Standard for testing React components by user behavior, not implementation details. |
| @testing-library/jest-dom | >=6.9.1 (dev) | DOM assertion matchers | Adds `.toBeInTheDocument()`, `.toHaveTextContent()`, etc. |
| @testing-library/user-event | >=14.6.1 (dev) | Simulated user interactions | Fires events the way a real user would (click, type, etc). |
| jsdom | >=29.0.2 (dev) | DOM environment | Provides browser-like DOM for headless testing. |

**Confidence:** HIGH -- Vitest is the standard for Vite projects. React Testing Library is the standard for React. These are uncontroversial choices.

**Configuration (`frontend/vite.config.js` addition):**
```js
// Add to existing vite.config.js
export default defineConfig({
  // ... existing config ...
  test: {
    environment: 'jsdom',
    globals: true,
    setupFiles: './src/test-setup.js',
  },
})
```

**What to test:** The bridge abstraction layer (`bridge.js`) -- mock `window.qt.webChannelTransport` and verify the JS-to-Python call contract. Settings panels -- verify form controls render and dispatch bridge calls. Theme switching -- verify CSS class application.

**What NOT to test:** QWebChannel internals, Vite build output, visual styling. These are either Qt-owned or better caught by manual testing.

**What NOT to use:**
- **Jest** -- Requires separate Babel/SWC config to handle JSX. Vitest uses Vite's existing transform pipeline. No reason to add a second build tool.
- **Cypress/Playwright for frontend** -- The React UI runs inside QWebEngineView, not a browser. Browser E2E tools cannot drive it. If E2E testing is ever needed, it would need to drive the Python app process and interact via QWebChannel mocking.

---

## Pre-commit Hooks

### pre-commit Framework

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| pre-commit | >=4.6.0 | Git hook management | Runs linters/formatters automatically before commits. Prevents dirty code from entering the repo. |

**Confidence:** HIGH -- pre-commit 4.6.0 released 2026-04-21. Universal standard for Python projects.

**Configuration (`.pre-commit-config.yaml`):**
```yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.15.11
    hooks:
      - id: ruff-check
        args: [--fix]
      - id: ruff-format
  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.20.2
    hooks:
      - id: mypy
        additional_dependencies: []
  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v5.0.0
    hooks:
      - id: trailing-whitespace
      - id: end-of-file-fixer
      - id: check-yaml
      - id: check-added-large-files
```

---

## Build Pipeline

### PyInstaller (App Bundling)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| PyInstaller | >=6.20.0 | Bundle Python app into standalone .exe | Already in use. v6.20.0 released 2026-04-22. Supports PySide6 out of the box. |
| pyinstaller-hooks-contrib | >=2026.4 | Community hooks for PyInstaller | Provides additional hooks for dependencies like keyboard, comtypes. |

**Confidence:** HIGH -- PyInstaller is the established choice for this project, and the latest version has excellent PySide6 support.

**Key fix needed:** The spec file is named `Windows Toolbox.spec` and must be renamed to `virelo.spec`. The spec must be updated to run `npm ci && npm run build` in the frontend directory (or the build script must do this before invoking PyInstaller).

### Inno Setup (Windows Installer)

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| Inno Setup | 6.7.1 | Windows installer builder | Already in use. v6.7.1 released 2026-02-24. Mature, free, well-documented. Dark mode support added in 6.6.0. |

**Confidence:** HIGH -- Inno Setup is the right tool. v7.0 is in preview but stick with 6.7.x for stability.

### Build Orchestration Script

The current `scripts/build-installer.ps1` needs to become a complete pipeline:

```
1. npm ci                           (in frontend/)
2. npm run build                    (produces frontend/dist/)
3. ruff check .                     (lint gate)
4. mypy .                           (type check gate)
5. pytest                           (test gate)
6. pyinstaller --clean virelo.spec  (bundle)
7. ISCC.exe installer/virelo.iss    (installer)
```

Steps 1-5 are quality gates. If any fails, the build stops. This is the contract that CI enforces.

---

## CI/CD

### GitHub Actions

| Technology | Version | Purpose | Why |
|------------|---------|---------|-----|
| GitHub Actions | N/A | CI pipeline | Free for public repos. Windows runners (`windows-latest`) available. No self-hosted runner needed. |

**Confidence:** HIGH -- GitHub Actions with `windows-latest` runner is the standard for open-source Windows desktop projects.

**Workflow structure (`.github/workflows/ci.yml`):**

```yaml
name: CI
on: [push, pull_request]

jobs:
  lint:
    runs-on: ubuntu-latest  # Linting doesn't need Windows
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.12' }
      - run: pip install ruff mypy
      - run: ruff check .
      - run: ruff format --check .
      - run: mypy .

  test-python:
    runs-on: windows-latest  # Tests need Win32 APIs
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.12' }
      - run: pip install -e ".[test]"
      - run: pytest

  test-frontend:
    runs-on: ubuntu-latest  # Frontend tests don't need Windows
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: '22' }
      - working-directory: frontend
        run: npm ci
      - working-directory: frontend
        run: npx vitest run
      - working-directory: frontend
        run: npx biome check .

  build:
    needs: [lint, test-python, test-frontend]
    runs-on: windows-latest  # Build must be Windows
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.12' }
      - uses: actions/setup-node@v4
        with: { node-version: '22' }
      - run: pip install -e ".[build]"
      - working-directory: frontend
        run: npm ci && npm run build
      - run: pyinstaller --clean --noconfirm virelo.spec
      - uses: actions/upload-artifact@v4
        with:
          name: virelo-build
          path: dist/Virelo/
```

**Key decisions:**
- Lint on Ubuntu (fast, cheap) -- ruff/mypy don't need Windows
- Python tests on Windows -- pywin32/comtypes need real Win32 APIs (some tests can mock this, but integration tests need it)
- Frontend tests on Ubuntu -- jsdom is platform-independent
- Build on Windows only -- PyInstaller cross-compilation is not supported
- Skip Inno Setup in CI unless doing a release (ISCC.exe must be installed on the runner)

---

## WebEngine Hardening Patterns

These are not library choices but code patterns for the existing PySide6/QWebEngineView stack.

### Block External Navigation

**Pattern:** Override `acceptNavigationRequest()` on the custom `QWebEnginePage` subclass.

```python
def acceptNavigationRequest(self, url, nav_type, is_main_frame):
    if url.scheme() in ("file", "qrc"):
        return True
    if url.scheme() == "http" and url.host() == "localhost":
        return True  # Allow Vite dev server
    logger.warning("Blocked navigation to: %s", url.toString())
    return False
```

**Confidence:** HIGH -- `acceptNavigationRequest` is the documented Qt API for this purpose. Verified in Qt 6.8+ docs.

### Disable Unnecessary WebEngine Features

```python
from PySide6.QtWebEngineCore import QWebEngineSettings

settings = web_view.settings()
settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanOpenWindows, False)
settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanAccessClipboard, False)
settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, False)
settings.setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, False)
```

**Confidence:** HIGH -- These are standard QWebEngineSettings attributes documented in Qt.

### Make Dev Mode Explicit

The existing `VIRELO_DEV` env var approach is correct. Harden by:
1. Never falling back to dev server URL if `VIRELO_DEV` is not set
2. Logging clearly when dev mode is active
3. Refusing to start if `frontend/dist/index.html` is missing in production mode (rather than showing a blank window)

### Request Interception (Defense in Depth)

```python
from PySide6.QtWebEngineCore import QWebEngineUrlRequestInterceptor

class VireloRequestInterceptor(QWebEngineUrlRequestInterceptor):
    def interceptRequest(self, info):
        url = info.requestUrl()
        if url.scheme() not in ("file", "qrc", "data"):
            if not (url.scheme() == "http" and url.host() == "localhost"):
                info.block(True)
```

**Confidence:** HIGH -- `QWebEngineUrlRequestInterceptor` is the documented API for request-level blocking.

---

## Version Management

### Single Source of Truth

Currently the version `1.4.2` appears in at least:
- `frontend/package.json`
- `installer/virelo.iss`
- Potentially `Windows Toolbox.spec`
- Possibly `app_config.py`

**Solution:** Define version once in `pyproject.toml`, read it at runtime via `importlib.metadata.version("virelo")`, and inject it into the frontend build and installer from there.

```toml
# pyproject.toml
[project]
name = "virelo"
version = "1.5.0"
```

```python
# virelo/__init__.py
from importlib.metadata import version, PackageNotFoundError
try:
    __version__ = version("virelo")
except PackageNotFoundError:
    __version__ = "dev"
```

**Confidence:** HIGH -- This is the official Python packaging recommendation (PEP 621 + importlib.metadata).

---

## Alternatives Considered

| Category | Recommended | Alternative | Why Not |
|----------|-------------|-------------|---------|
| Python linter | ruff | flake8/pylint | Ruff is 10-100x faster and replaces both plus isort/black/pyupgrade |
| Python formatter | ruff format | black | ruff format produces identical output, no separate tool needed |
| Python type checker | mypy | pyright, ty | mypy has best pre-commit/CI integration; ty too new (preview) |
| JS linter/formatter | biome | ESLint 10 + Prettier | ESLint 10 broke eslint-plugin-react; Biome is faster and simpler for a small JSX project |
| JS test runner | vitest | Jest | Vitest shares Vite config, Jest would need separate Babel setup |
| Python test runner | pytest | unittest | pytest has richer fixtures, plugins (pytest-qt, pytest-cov), and cleaner syntax |
| CI platform | GitHub Actions | Azure Pipelines, Jenkins | GitHub Actions is free, integrated, has Windows runners, zero infrastructure |
| Bundler | PyInstaller | Nuitka, cx_Freeze | PyInstaller is already working and has the best PySide6 support |
| Installer | Inno Setup 6.7 | NSIS, WiX | Inno Setup is already working, free, has dark mode, simplest scripting |
| Pre-commit | pre-commit | husky (JS) | pre-commit handles both Python and JS hooks from one config file |

---

## Installation Commands

### Python Development Dependencies

```bash
# After pyproject.toml migration, install with:
pip install -e ".[dev]"

# Which installs: ruff, mypy, pytest, pytest-cov, pytest-qt, pre-commit
```

**pyproject.toml dependency groups:**
```toml
[project.optional-dependencies]
dev = [
    "ruff>=0.15.11",
    "mypy>=1.20.2",
    "pytest>=9.0.2",
    "pytest-cov>=7.1.0",
    "pytest-qt>=4.5.0",
    "pre-commit>=4.6.0",
]
build = [
    "pyinstaller>=6.20.0",
    "pyinstaller-hooks-contrib>=2026.4",
]
```

### Frontend Development Dependencies

```bash
cd frontend
npm install -D vitest @testing-library/react @testing-library/jest-dom @testing-library/user-event jsdom @biomejs/biome
```

---

## Pinned Existing Stack (No Changes)

These are already in use and should stay at their current constraints. Listed for completeness.

| Technology | Current Constraint | Latest Available | Action |
|------------|-------------------|------------------|--------|
| PySide6 | >=6.6 | 6.11.0 | Widen to >=6.8 (drops Python 3.8/3.9 which are EOL) |
| React | ^19.1.0 | 19.1.0 | No change needed |
| Vite | ^6.3.4 | 6.3.4 | No change needed |
| @vitejs/plugin-react | ^4.5.2 | 4.5.2 | No change needed |
| keyboard | >=0.13.5 | 0.13.5 | No change |
| pywin32 | >=306 | 306+ | No change |
| comtypes | >=1.3.0 | 1.3.0+ | No change |

---

## Sources

- [pytest 9.0.2 on PyPI](https://pypi.org/project/pytest/) -- HIGH confidence
- [ruff 0.15.11 on PyPI](https://pypi.org/project/ruff/) -- HIGH confidence
- [Ruff v0.15.0 announcement](https://astral.sh/blog/ruff-v0.15.0) -- HIGH confidence
- [mypy 1.20.2 on PyPI](https://pypi.org/project/mypy/) -- HIGH confidence
- [mypy 1.20 blog post](https://mypy-lang.blogspot.com/2026/03/mypy-120-released.html) -- HIGH confidence
- [Vitest on npm](https://www.npmjs.com/package/vitest) -- HIGH confidence
- [React Testing Library on npm](https://www.npmjs.com/package/@testing-library/react) -- HIGH confidence
- [PyInstaller 6.20.0 docs](https://pyinstaller.org/en/stable/) -- HIGH confidence
- [Inno Setup 6.7.1 downloads](https://jrsoftware.org/isdl.php) -- HIGH confidence
- [pre-commit 4.6.0 on PyPI](https://pypi.org/project/pre-commit/) -- HIGH confidence
- [pytest-qt on PyPI](https://pypi.org/project/pytest-qt/) -- HIGH confidence
- [pytest-cov 7.1.0 on PyPI](https://pypi.org/project/pytest-cov/) -- HIGH confidence
- [Biome v2 docs](https://biomejs.dev/) -- MEDIUM-HIGH confidence
- [ESLint v10 / eslint-plugin-react compatibility issue](https://github.com/jsx-eslint/eslint-plugin-react/issues/3977) -- HIGH confidence
- [QWebEnginePage acceptNavigationRequest docs](https://doc.qt.io/qtforpython-6.8/PySide6/QtWebEngineCore/QWebEnginePage.html) -- HIGH confidence
- [QWebEngineUrlRequestInterceptor docs](https://doc.qt.io/qtforpython-6/PySide6/QtWebEngineCore/QWebEngineUrlRequestInterceptor.html) -- HIGH confidence
- [PySide6 6.11.0 on PyPI](https://pypi.org/project/PySide6/) -- HIGH confidence
- [Python single-source version guide](https://packaging.python.org/en/latest/discussions/single-source-version/) -- HIGH confidence
- [pyinstaller-hooks-contrib 2026.4 on PyPI](https://pypi.org/project/pyinstaller-hooks-contrib/) -- HIGH confidence

---

*Stack research: 2026-04-24*
