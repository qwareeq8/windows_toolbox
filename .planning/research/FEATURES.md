# Feature Landscape

**Domain:** Windows desktop utility -- release hygiene, build infrastructure, code hardening
**Researched:** 2026-04-24
**Focus:** Infrastructure/hygiene features for a personal-use app heading to GitHub

## Table Stakes

Features that must exist for a credible GitHub repository. Missing any of these signals "unfinished side project" rather than "presentable tool."

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| `.gitignore` | Prevents committing `__pycache__/`, `dist/`, `build/`, `frontend/node_modules/`, `frontend/dist/`, `.env`, `*.pyc`, `.venv/` | Low | No .gitignore exists today. Use GitHub's Python template as base, add Node, PyInstaller, Inno Setup patterns. |
| Remove stale "Windows Toolbox" naming | Spec file is literally named `Windows Toolbox.spec`; build script references it. Product is Virelo. | Low | Rename spec to `virelo.spec`, update `build-installer.ps1` references. Three stale references remain. |
| Single source of truth for version | Version `1.4.2` is hardcoded in 5+ locations: `virelo.iss`, `package.json`, `panels.jsx`, `pages.jsx`, `app.jsx`. Any version bump requires editing all of them. | Medium | Create `app_config.py:VERSION`, read it in the Inno Setup script via `/D` flag, inject into frontend at build time via Vite `define`. The `package.json` version is independent (frontend is private, never published). |
| Working build pipeline from clean checkout | Current pipeline is incomplete: `build-installer.ps1` runs PyInstaller but never runs `npm ci` or `npm run build` first. The PyInstaller spec requires `frontend/dist/` to exist. A clean checkout fails. | Medium | Add `npm ci && npm run build` step before PyInstaller in the build script. Validate `frontend/dist/index.html` exists before proceeding. |
| README with install/usage/build instructions | GitHub surfaces README automatically. Without one, visitors cannot understand what the app does or how to build it. | Low | Sections: What it is, screenshot, install (download installer), build from source, usage, license. |
| LICENSE file | GitHub expects it. Without one, the code has no explicit license, which means "all rights reserved" by default. | Low | MIT is standard for personal utilities. Add `LICENSE` in root. |
| Remove fake/no-op UI controls | "Automatic updates" toggle, "Anonymous telemetry" toggle, and "Start minimized to tray" toggle exist in `pages.jsx` with state in `app.jsx` but connect to nothing in the backend. They mislead users. | Low | Delete the three fake `<Row>` elements from `GeneralPage`. Remove `startTray`, `autoUpdate`, `telemetry` from React initial state. |
| WebEngine: block external navigation | `VireloWebPage` does not override `acceptNavigationRequest`. Any link in the frontend (or injected content) can navigate the embedded browser to any URL. | Low | Override `acceptNavigationRequest` in `VireloWebPage` to only allow `file://` and `qrc://` schemes (plus `localhost` in dev mode). Redirect external URLs to `QDesktopServices.openUrl()`. Qt's own examples demonstrate this exact pattern. |
| WebEngine: disable unnecessary permissions | `LocalContentCanAccessRemoteUrls` is set to `True` in `webview.py`. This allows the local frontend to fetch arbitrary remote URLs, which is unnecessary for a settings UI that only talks to the Python bridge via QWebChannel. | Low | Set `LocalContentCanAccessRemoteUrls` to `False`. The QWebChannel bridge does not use HTTP -- it uses Qt's internal IPC. Only re-enable in dev mode if needed for Vite HMR. |
| Handle missing frontend build gracefully | If `frontend/dist/` is missing in a release build, `_get_frontend_url` falls back to a path that may not exist and loads a blank page with no error message. | Low | Show a clear error message (QMessageBox or styled HTML) when `frontend/dist/index.html` is not found. Log the expected path. |
| `CLAUDE.md` project configuration | Standard for Claude Code projects. Documents build commands, conventions, project layout, and rules. | Low | Add at project root. Include: how to build, how to run dev mode, naming conventions, file structure overview, testing commands. Under 200 lines. |

## Differentiators

Features that elevate Virelo from "works on my machine" to "well-engineered personal tool." Not expected by casual visitors, but valued by anyone reading the code.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| Structured state bridge with draft model | Current bridge saves settings immediately on every `save_settings()` call. No draft/discard/cancel flow. The React side tracks `unsaved` state but the Python side has no concept of draft vs committed. A proper draft model lets users preview changes and discard them. | Medium | Add a `DraftSettingsState` class that holds uncommitted changes. Bridge `save_settings()` writes to draft first, `commit_settings()` persists to QSettings. `discard_settings()` reverts draft to committed. React already has the unsaved indicator. |
| Python package structure (`virelo/`) | All 12 Python files sit flat in the project root. No `__init__.py`, no package namespace. This makes imports fragile, testing harder, and PyInstaller hiddenimports necessary. | Medium | Move Python source into `virelo/` package: `virelo/__init__.py`, `virelo/main.py`, `virelo/bridge.py`, etc. Update the PyInstaller spec entry point. Remove explicit `hiddenimports` that are only needed because of the flat layout. |
| Split `main.py` into focused modules | `main.py` is 1451 lines containing the entry point, MainWindow, ShiftSnapRestore engine, Win32 helpers, startup shortcut logic, and ~30 utility functions. | Medium | Extract: `virelo/win32_helpers.py`, `virelo/snap_restore.py`, `virelo/startup.py`. Keep MainWindow and entry point in `main.py`. CONCERNS.md already has a precise extraction plan. |
| Tests for pure-logic modules | Zero test files exist. No `conftest.py`, no `pytest.ini`. The most testable code (SettingsState validation, ExplorerAutosizeEngine state machine, path canonicalization, theme resolution) has no coverage. | Medium | Add `tests/` directory with `pytest` configuration. Start with `test_settings_state.py`, `test_theme.py`, `test_app_config.py` -- pure functions that need no GUI or COM mocking. Target the highest-risk validation code first. |
| Ruff linting and formatting | No linter configured. No formatting standard enforced. Code style varies across files. | Low | Add `ruff` to dev dependencies. Create `pyproject.toml` with `[tool.ruff]` section. Configure line length, target Python version, select rule sets (E, F, W, I for imports). Add `ruff check` and `ruff format --check` to CI. |
| GitHub Actions CI pipeline | No CI exists. No automated build verification, no lint checks, no test runs. | Medium | Create `.github/workflows/ci.yml`: lint with ruff, run pytest, build frontend, run PyInstaller. Use `windows-latest` runner (required for PySide6/Win32 testing). Inno Setup is pre-installed on GitHub's Windows runners. |
| Module registry architecture | PROJECT.md mentions preparing a registry for future personal tools. This is forward-looking infrastructure: a `ModuleInfo` dataclass, a registry class that discovers installed modules, and a sidebar that adapts to registered modules. | High | Define `virelo/registry.py` with `ModuleInfo(name, icon, page_component, bridge_slots)`. The registry enumerates available modules at startup. Snap and Explorer become registered modules. This enables future tools without a plugin system. |
| Consolidate duplicate code | Three path canonicalization implementations, two `resource_path` copies, two tab state dataclasses, two identical autosize functions. | Low-Med | Follow the precise deduplication plan in CONCERNS.md. Single `canonicalize_path()` in a shared utility, single `resource_path()` in `app_config.py` or a new `paths.py`, remove dead `TabState`, merge identical autosize functions. |
| Remove deprecated Qt attributes | `main.py` calls `setAttribute(AA_EnableHighDpiScaling)` and `setAttribute(AA_UseHighDpiPixmaps)` which are no-ops in Qt 6 and generate deprecation warnings. | Low | Delete the two `setAttribute` calls. They are always enabled in Qt 6 and the calls do nothing except produce log noise. |
| `pyproject.toml` for project metadata | No `pyproject.toml` exists. Version, dependencies, and tool config are scattered across `requirements.txt`, `app_config.py`, and nowhere for linting/testing. | Low | Create `pyproject.toml` with `[project]` metadata (name, version, description, requires-python), `[tool.ruff]` config, `[tool.pytest.ini_options]` config. Keep `requirements.txt` for pip install convenience but `pyproject.toml` becomes the source of truth. |

## Anti-Features

Features to explicitly NOT build. These are either already called out as out-of-scope in PROJECT.md, or represent common over-engineering traps for personal utilities.

| Anti-Feature | Why Avoid | What to Do Instead |
|--------------|-----------|-------------------|
| Plugin system / third-party extensibility | Personal-use tool. The internal module registry provides enough structure for adding personal tools. A plugin API adds security surface, API stability burden, and documentation overhead for zero users. | Use the internal module registry pattern. New tools are added by the developer, not discovered/loaded at runtime. |
| Auto-update mechanism | No distribution infrastructure exists. No update server, no code signing certificate, no delta update logic. The fake "Automatic updates" toggle in the UI is actively misleading. | Remove the fake toggle. Users update by downloading a new installer from GitHub Releases. Document this in the README. |
| Telemetry / analytics | Personal tool used by one person. Telemetry adds privacy concerns, network requests, data storage, and opt-in/out UX for no benefit. | Remove the fake "Anonymous telemetry" toggle. If crash diagnostics are needed, the existing `crash.log` and `virelo.log` files are sufficient. |
| Cross-platform support | Deep Win32/COM dependency throughout. The app uses `ctypes` for `USER32.dll`, `DWMAPI.dll`, COM automation for Explorer, Windows keyboard hooks, and Windows registry for settings. Porting is a rewrite, not a feature. | Keep Windows-only. Document the platform requirement clearly in README. |
| Code signing | Requires purchasing an EV certificate ($200-400/year), setting up HSM or cloud signing service, and maintaining signing infrastructure. For a personal-use tool with one user, the cost/benefit is negative. | Users will see a Windows SmartScreen warning on first run. Document this in the README ("Windows may show a warning because the app is not code-signed"). |
| MSI installer (instead of Inno Setup) | Inno Setup works, is already configured, and produces a professional installer. MSI adds complexity (WiX toolset, XML manifests) for no user-facing benefit. | Keep Inno Setup. It is well-supported, has a modern wizard style, and integrates with GitHub Actions runners. |
| Web-based documentation site | A README and inline code comments are sufficient for a personal tool. A dedicated docs site (Docusaurus, MkDocs) adds build steps, hosting, and maintenance for one user. | README.md + CLAUDE.md + inline comments cover all documentation needs. |
| Internationalization (i18n) | Single-user, English-only tool. Adding i18n infrastructure (string extraction, translation files, locale detection) adds complexity to every UI string for zero benefit. | Hardcode English strings. |
| Custom crash reporting service | The app already has `faulthandler` writing to `crash.log` and a rotating file logger. Sentry or similar services add network dependencies and privacy concerns. | Keep local `crash.log` + `virelo.log`. These are sufficient for debugging a personal tool. |

## Feature Dependencies

```
.gitignore ------> (independent, do first)
Remove stale naming ------> (independent, do first)
Remove fake UI controls ------> (independent, do first)
Remove deprecated Qt attrs ------> (independent, do first)

LICENSE + README ------> (independent, but README benefits from having build pipeline working)

Working build pipeline ------> Single source version (version needs to flow through build)
                       ------> Handle missing frontend build (build validates this)

WebEngine navigation block ------> (independent)
WebEngine disable remote URLs ------> (independent)

pyproject.toml ------> Ruff config (ruff config lives in pyproject.toml)
               ------> pytest config (pytest config lives in pyproject.toml)

Ruff linting ------> pyproject.toml (config home)
Tests ------> pyproject.toml (config home)
      ------> Python package structure (imports cleaner as a package)

Python package structure ------> Split main.py (easier to reorganize into package)
                         ------> Consolidate duplicates (dedup during reorganization)

GitHub Actions CI ------> Working build pipeline (CI needs build to work)
                  ------> Ruff linting (CI runs lint)
                  ------> Tests (CI runs tests)

Structured state bridge ------> Split main.py (bridge refactoring needs cleaner code structure)

Module registry ------> Python package structure (registry is a package-level concern)
                ------> Structured state bridge (modules need a clean bridge pattern)

CLAUDE.md ------> (independent, but benefits from settled project structure)
```

## MVP Recommendation

For making the repo presentable on GitHub, prioritize in this order:

1. **Repository hygiene (table stakes, low complexity):**
   - `.gitignore` -- prevents accidental commits of build artifacts
   - Remove stale "Windows Toolbox" naming -- rename spec file, update build script
   - Remove fake UI controls -- delete misleading toggles for auto-update, telemetry, start-minimized
   - Remove deprecated Qt attributes -- eliminate log noise
   - `LICENSE` file -- MIT license in project root
   - `CLAUDE.md` -- project configuration for Claude Code

2. **Build reliability (table stakes, medium complexity):**
   - Working build pipeline -- add npm build step to `build-installer.ps1`
   - Single source of version truth -- `app_config.py:VERSION` as canonical, inject elsewhere
   - Handle missing frontend build -- error message instead of blank page

3. **Security hardening (table stakes, low complexity):**
   - WebEngine navigation blocking -- override `acceptNavigationRequest`
   - WebEngine disable remote URLs -- set `LocalContentCanAccessRemoteUrls` to `False`

4. **Code quality infrastructure (differentiator, medium complexity):**
   - `pyproject.toml` -- project metadata and tool configuration home
   - Ruff linting -- automated code quality
   - Tests for pure-logic modules -- regression safety net
   - `README.md` -- install/build/usage documentation

5. **Code structure (differentiator, medium complexity):**
   - Consolidate duplicate code -- reduce maintenance burden
   - Split `main.py` -- focused modules
   - Python package structure -- proper `virelo/` namespace

**Defer:**
- Module registry: High complexity, forward-looking. Only valuable when a second tool is actually being added. Building it now is speculative engineering.
- Structured state bridge: Medium complexity. The current save-immediately pattern works. Draft state is a UX improvement but not blocking for GitHub release.
- GitHub Actions CI: Depends on build pipeline, linting, and tests all working first. Add after those are stable.

## Sources

- [GitHub repository best practices](https://docs.github.com/en/repositories/creating-and-managing-repositories/best-practices-for-repositories) -- HIGH confidence
- [Qt for Python QWebEnginePage.acceptNavigationRequest](https://doc.qt.io/qtforpython-6/PySide6/QtWebEngineCore/QWebEnginePage.html) -- HIGH confidence (Context7 verified, official Qt examples confirm pattern)
- [Qt Markdown Editor example -- PreviewPage navigation control](https://doc.qt.io/qtforpython-6/examples/example_webenginewidgets_markdowneditor.html) -- HIGH confidence (official Qt example)
- [Python Packaging User Guide -- single-sourcing version](https://packaging.python.org/en/latest/discussions/single-source-version/) -- HIGH confidence
- [Ruff pre-commit integration](https://github.com/astral-sh/ruff-pre-commit) -- HIGH confidence
- [GitHub Actions Python CI](https://docs.github.com/en/actions/tutorials/build-and-test-code/python) -- HIGH confidence
- [PyInstaller documentation](https://pyinstaller.org/en/stable/spec-files.html) -- HIGH confidence
- [GitHub gitignore Python template](https://github.com/github/gitignore/blob/main/Python.gitignore) -- HIGH confidence
- [Draft state UX patterns -- Cloudscape Design System](https://cloudscape.design/patterns/general/unsaved-changes/) -- MEDIUM confidence
- [CLAUDE.md best practices](https://code.claude.com/docs/en/memory) -- HIGH confidence
