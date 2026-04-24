# Project Research Summary

**Project:** Virelo (PySide6 + React + QWebEngineView hybrid desktop utility)
**Domain:** Windows desktop utility hardening and refactoring -- cleanup, build infrastructure, code quality
**Researched:** 2026-04-24
**Confidence:** HIGH

## Executive Summary

Virelo is a personal Windows desktop utility transitioning from a working prototype to a presentable, maintainable codebase. The core application stack (PySide6, React 19, Vite, QWebChannel, PyInstaller, Inno Setup) is established and correct -- the research focus was entirely on the tooling, hardening, and structural changes required to make the codebase defensible and the build reproducible. The recommended approach is a bottom-up refactoring sequence: fix the build pipeline and security surface first (both have immediate risk), then add quality infrastructure (linting, tests, CI), then restructure the code (package layout, module splitting). This sequence avoids the most dangerous failure mode, which is attempting to restructure a monolith before the build is reliable enough to catch regressions.

The most significant risk is the interaction between two categories of pitfalls: PyInstaller build failures caused by a missing frontend build step (currently guaranteed to fail from a clean checkout), and the security exposure of an unrestricted QWebEngine running in an admin-elevated process. Both must be addressed before the repository is made public. The secondary risk category is the refactoring sequence itself -- splitting main.py into a proper virelo/ package must proceed bottom-up (pure functions first, Qt classes last) to avoid circular imports, signal/slot disconnection bugs, and COM apartment threading violations that are extremely difficult to diagnose.

The feature scope is clear and bounded. There are eleven active requirements in PROJECT.md, all of which fall into well-understood categories: hygiene (naming, gitignore, license), build reliability, WebEngine hardening, code quality infrastructure, and structural refactoring. Nothing in the active requirements requires novel technology research -- the patterns are documented and the tooling choices have high-confidence backing. Three categories of features (plugin system, auto-update, telemetry) are explicitly anti-features and must be removed from the fake UI controls that currently mislead users.

---

## Key Findings

### Recommended Stack

The existing application stack needs no changes. The research identified the tooling layer to add on top: Ruff (>=0.15.11) replaces all Python linting and formatting tools in a single dependency; mypy (>=1.20.2) provides gradual type checking starting with permissive settings; pytest (>=9.0.2) with pytest-qt and pytest-cov provides the test framework; Biome (>=2.3) replaces both ESLint and Prettier for the React frontend (ESLint 10 broke eslint-plugin-react compatibility as of April 2026); Vitest (>=4.1.4) shares the Vite config for frontend tests with zero additional configuration; pre-commit (>=4.6.0) enforces quality gates at commit time; GitHub Actions with windows-latest runners provides free CI for a Windows-only project.

The version pinning recommendation is Python 3.12 (most stable with PySide6 + PyInstaller), PySide6 widened to >=6.8 (drops EOL Python 3.8/3.9), and PyInstaller pinned to >=6.20.0 (latest, with the best QtWebEngine hook support). Version 1.4.2 must be consolidated from 5+ locations into a single pyproject.toml source, read at runtime via importlib.metadata.

**Core technologies (tooling layer to add):**
- **ruff >=0.15.11**: Python linting + formatting -- replaces flake8, black, isort, pyupgrade in one tool
- **mypy >=1.20.2**: Static type checking -- critical for Win32/COM code where wrong types cause silent failures
- **pytest >=9.0.2 + pytest-qt + pytest-cov**: Python testing -- only framework with real PySide6 widget support
- **Biome >=2.3**: JS/JSX linting + formatting -- ESLint 10 broke eslint-plugin-react, Biome is faster and simpler
- **Vitest >=4.1.4 + @testing-library/react**: Frontend testing -- shares Vite config, zero extra bundler setup
- **pre-commit >=4.6.0**: Git hook management -- enforces ruff + mypy before every commit
- **GitHub Actions (windows-latest)**: CI -- free, has Windows runners, supports PyInstaller builds
- **pyproject.toml (PEP 621)**: Project configuration hub -- single source for metadata, version, all tool configs

### Expected Features

Features are cleanly divided between work that must happen before GitHub publication (table stakes) and work that elevates the codebase quality (differentiators). Three features are explicitly anti-features that must be removed, not deferred.

**Must have (table stakes -- without these the repo signals unfinished):**
- `.gitignore` -- no gitignore exists; build artifacts, node_modules, .venv will be committed on first push
- Remove stale "Windows Toolbox" naming -- spec file literally named `Windows Toolbox.spec`, build scripts reference it
- Working build pipeline from clean checkout -- currently fails: npm build step missing before PyInstaller
- Single source of truth for version -- version 1.4.2 duplicated in 5+ files
- Handle missing frontend build gracefully -- blank page instead of error when `frontend/dist/` missing
- WebEngine navigation blocking -- no `acceptNavigationRequest` override; admin-elevated app can navigate to internet
- WebEngine disable remote URLs -- `LocalContentCanAccessRemoteUrls = True` in production is unnecessary and dangerous
- Remove fake UI controls -- auto-update, telemetry, start-minimized toggles connect to nothing
- Remove deprecated Qt 6 attributes -- `AA_EnableHighDpiScaling` / `AA_UseHighDpiPixmaps` are no-ops, generate warnings
- LICENSE file -- MIT, required for any public repo
- README with install/usage/build instructions -- no README exists
- CLAUDE.md project configuration

**Should have (differentiators -- elevate from works to well-engineered):**
- pyproject.toml with project metadata and tool configuration
- Ruff linting and formatting
- Tests for pure-logic modules (SettingsState, theme resolution, app_config)
- Python package structure (`virelo/` namespace)
- Split `main.py` into focused modules (platform, services, workers, app, bridge)
- Consolidate duplicate code (three path canonicalization implementations, two `resource_path` copies)
- GitHub Actions CI pipeline
- Structured state bridge with draft model

**Defer (not essential for first GitHub release):**
- Module registry architecture -- high complexity, speculative until a second tool is actually being added
- Structured state bridge -- medium complexity; save-immediately pattern works for now
- GitHub Actions CI -- add after build pipeline, linting, and tests are stable

**Remove (anti-features -- delete, not defer):**
- Auto-update toggle and backend stubs
- Telemetry toggle
- Start-minimized-to-tray toggle (currently connects to nothing)

### Architecture Approach

The recommended architecture preserves the existing host-shell + embedded-SPA pattern (PySide6 hosts React in QWebEngineView via QWebChannel) and restructures the Python side from a flat-file monolith into a proper `virelo/` package with domain-oriented subpackages. The core principle is strict data flow: React never calls services directly, the bridge is the only crossing point, and services never import the bridge. MainWindow becomes a wiring hub that creates and connects components rather than owning business logic. The refactoring must proceed bottom-up: pure leaf modules first (platform helpers, settings), then services, then bridge and app, then entry point cleanup.

**Major components:**
1. `virelo/platform/` -- pure Win32/ctypes helpers, leaf module, no virelo imports; independently testable
2. `virelo/settings/` -- Settings (QSettings), SettingsState (JSON facade + validation), config (constants)
3. `virelo/services/` -- SnapService, ExplorerService, ThemeService, StartupService; own their worker lifecycles
4. `virelo/workers/` -- KeyCaptureWorker, ExplorerAutosizeWorker; QThread-aware, no service imports
5. `virelo/bridge.py` -- VireloBridge QObject; only Slots/Signals, calls services (not MainWindow)
6. `virelo/app.py` -- MainWindow as wiring hub; creates, wires, and starts all components
7. `virelo/webview.py` -- VireloWebView + VireloWebPage; hardened navigation, channel setup
8. `frontend/src/` -- React SPA; no structural changes needed during this refactoring milestone

### Critical Pitfalls

1. **PyInstaller builds without frontend assets** -- The build script never runs `npm ci && npm run build` before PyInstaller. Fix: add npm build step first; assert `frontend/dist/index.html` exists before PyInstaller runs. This is guaranteed to fail from a clean checkout today.

2. **QWebEngine navigation unrestricted in admin-elevated process** -- No `acceptNavigationRequest` override; `LocalContentCanAccessRemoteUrls = True` in production. Fix: override `acceptNavigationRequest` to allow only `file://` and dev `localhost`; set `LocalContentCanAccessRemoteUrls = False` in non-dev mode. A compromised npm dep or XSS could call bridge slots as administrator.

3. **Splitting main.py causes signal/slot disconnection and circular imports** -- Monolith refactoring must be bottom-up. Fix: extract pure functions first (win32 helpers, path utils), then services, then Qt classes last. Use explicit `parent=` on all QObjects. Test full lifecycle after each extraction.

4. **COM apartment threading violations during refactoring** -- ExplorerAutosizeWorker initializes COM as STA; moving any COM call to a different thread causes hard crashes. Fix: keep COM init + use + uninit co-located in the same thread; document the threading contract; never pass COM objects across thread boundaries.

5. **PyInstaller + QWebEngine resource collection failures** -- QtWebEngineProcess.exe and .pak resource files must be in specific relative paths; hooks change between versions. Fix: pin both PyInstaller and PySide6 versions; verify `dist/Virelo/` contains `QtWebEngineProcess.exe`, `resources/`, and `translations/` after every build.

---

## Implications for Roadmap

Based on combined research, the active requirements map cleanly to six phases. The ordering is dictated by: (1) security and build reliability must precede any public exposure, (2) quality infrastructure must precede structural refactoring (need tests to catch regressions), (3) structural refactoring must proceed bottom-up per the dependency graph.

### Phase 1: Identity and Hygiene

**Rationale:** Zero-risk, independent changes that unblock everything else. Stale "Windows Toolbox" naming in the spec file will break the renamed build pipeline. No .gitignore means the first `git add .` commits 50+ unwanted files.

**Delivers:** Clean repo state -- correct naming throughout, no accidental build artifact commits, MIT license, basic README skeleton.

**Addresses:** Remove stale naming, .gitignore, LICENSE, CLAUDE.md, remove deprecated Qt attributes, remove fake UI controls.

**Avoids:** Pitfall 15 (spec file with spaces), Pitfall 13 (deprecated Qt attributes), Pitfall 8 (fake control removal leaving orphaned state).

### Phase 2: Build Pipeline and Security Hardening

**Rationale:** The two must-fix-before-going-public changes. The build is currently broken from a clean checkout (missing npm step). The WebEngine runs in an admin process with no navigation restrictions. Both require touching `build-installer.ps1`, `webview.py`, and the PyInstaller spec -- the same files, at the same time.

**Delivers:** Reproducible build from clean checkout; hardened WebEngine host that cannot navigate to external URLs; single version source in `pyproject.toml`.

**Addresses:** Working build pipeline, WebEngine navigation blocking, WebEngine remote URL disable, single source of version truth, handle missing frontend build gracefully.

**Avoids:** Pitfall 1 (PyInstaller without frontend assets), Pitfall 2 (navigation unrestricted), Pitfall 6 (WebEngine resource collection), Pitfall 7 (version duplication).

**Uses:** PyInstaller >=6.20.0, pyproject.toml (PEP 621), `acceptNavigationRequest` pattern, `QWebEngineUrlRequestInterceptor`.

### Phase 3: Code Quality Infrastructure

**Rationale:** Linting, formatting, and tests must exist before structural refactoring begins. The refactoring (Phase 4) will touch nearly every Python file; without a test suite, regressions are invisible. Ruff and mypy catch issues immediately. GitHub Actions CI closes the loop. This phase has no structural risk -- it only adds tooling, it does not move code.

**Delivers:** Green CI on every push; pytest suite covering pure-logic modules; ruff + mypy clean codebase; pre-commit hooks preventing dirty commits.

**Addresses:** pyproject.toml, Ruff linting, tests for pure-logic modules, pre-commit hooks, GitHub Actions CI.

**Avoids:** Pitfall 11 (CI admin/COM limitations -- tier tests correctly), Pitfall 4 (refactoring without tests).

**Uses:** ruff >=0.15.11, mypy >=1.20.2, pytest >=9.0.2 + pytest-qt + pytest-cov, Biome >=2.3, Vitest >=4.1.4, pre-commit >=4.6.0, GitHub Actions windows-latest.

### Phase 4: Module Splitting and Package Structure

**Rationale:** Highest-risk phase, gated on Phase 3 (tests). The 1451-line `main.py` monolith splits into a `virelo/` package following the bottom-up sequence: `platform/` first, then `settings/`, then `services/`, then `workers/`, then bridge and app. Each extraction step verified against the full lifecycle test.

**Delivers:** Maintainable `virelo/` package with clear module boundaries; MainWindow reduced to a wiring hub; bridge decoupled from MainWindow internals; duplicate code eliminated.

**Addresses:** Python package structure, split main.py, consolidate duplicate code, remove bridge-to-MainWindow coupling.

**Avoids:** Pitfall 4 (signal/slot disconnection), Pitfall 5 (COM apartment threading), Pitfall 9 (resource_path duplication).

### Phase 5: Snap Architecture and Explorer Feature Hardening

**Rationale:** After the package structure is in place, service-level improvements to snap and Explorer can be made safely in their isolated service modules.

**Delivers:** Snap service with keyboard hook race condition fix; Explorer service with COM cache resilience; both features with accurate UI scope.

**Addresses:** Improve snap architecture (hotkey/movement separation, exclusions, multi-monitor), Explorer feature scope cleanup.

**Avoids:** Pitfall 10 (keyboard hook race condition), Pitfall 12 (comtypes cache corruption).

### Phase 6: Repository Completion and Module Registry Groundwork

**Rationale:** Final polish and forward-looking infrastructure. The module registry is deferred to this phase because it is speculative until Phase 4 confirms the package structure.

**Delivers:** Fully presentable GitHub repository; module registry skeleton enabling future tools without structural changes.

**Addresses:** README with install/usage/build documentation, module registry architecture groundwork.

### Phase Ordering Rationale

- Phases 1-2 must precede all others: Phase 1 unblocks build, Phase 2 fixes build and security before public exposure.
- Phase 3 must precede Phase 4: cannot safely refactor a monolith without a test suite.
- Phase 4 must precede Phases 5-6: service improvements and registry require the package structure to be in place.
- Phase 5 and 6 can be reordered based on priority.
- The feature dependency graph from FEATURES.md confirms this ordering: pyproject.toml enables ruff and pytest; Python package structure enables module splitting and registry; CI requires working build, linting, and tests.

### Research Flags

Phases with well-documented patterns (skip research-phase):

- **Phase 1 (Identity/Hygiene):** Pure rename and file creation. No research needed.
- **Phase 2 (Build/Security):** `acceptNavigationRequest`, `QWebEngineUrlRequestInterceptor`, `pyproject.toml` -- all verified against official Qt and Python packaging docs.
- **Phase 3 (Code Quality):** Ruff, mypy, pytest, Biome, Vitest, GitHub Actions -- mainstream tools with official documentation and configuration examples in STACK.md.
- **Phase 6 (Repository Completion):** README/LICENSE -- standard practice, no research needed.

Phases that may benefit from targeted research during planning:

- **Phase 4 (Module Splitting):** The bottom-up sequence is well-reasoned but the specific signal/slot connection topology of `main.py` warrants a pre-phase codebase walkthrough to map every connection before moving any code. Internal codebase audit, not external research.
- **Phase 5 (Snap/Explorer Hardening):** Multi-monitor edge cases and the `keyboard` library hook lifecycle under rapid rebinding warrant a focused code review before the fix design is finalized.

---

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | All tools verified on PyPI/npm with exact versions. Biome rated MEDIUM-HIGH (v2.3 stable but newer than ESLint ecosystem). |
| Features | HIGH | Feature list derived directly from PROJECT.md active requirements and codebase analysis. No speculation. |
| Architecture | HIGH | Package layout and component boundaries follow established Python packaging and Qt architectural patterns, sourced from official docs. |
| Pitfalls | HIGH | Pitfalls 1-17 derived from direct codebase analysis combined with PyInstaller and Qt issue trackers. |

**Overall confidence:** HIGH

### Gaps to Address

- **mypy coverage ramp-up:** Codebase has zero type annotations. The plan for tightening mypy strictness over time is not specified. Recommend defining a per-phase strictness target (e.g., Phase 4 completion = typed public APIs on all service classes).

- **Smoke test design for PyInstaller output:** PITFALLS.md recommends a `--smoke-test` flag. The exact implementation (how to detect QWebChannel initialization success from the Python side without an interactive window) is not fully specified. Needs design during Phase 3 CI setup.

- **Biome version stability:** Biome v2.3 has a smaller community track record than ESLint. Pin to an exact version in `.pre-commit-config.yaml` to avoid unexpected breakage.

- **Admin elevation in CI:** GitHub Actions `windows-latest` runners do not run as administrator. Tests requiring admin (keyboard hooks, startup shortcut creation, some Win32 APIs) cannot run in standard CI. The test tiering (unit vs integration) must be designed explicitly in Phase 3.

---

## Sources

### Primary (HIGH confidence)

- [ruff 0.15.11 on PyPI](https://pypi.org/project/ruff/) -- linting/formatting tool selection
- [mypy 1.20.2 on PyPI](https://pypi.org/project/mypy/) + [mypy 1.20 blog](https://mypy-lang.blogspot.com/2026/03/mypy-120-released.html) -- type checking
- [pytest 9.0.2 on PyPI](https://pypi.org/project/pytest/) + [pytest-qt](https://pypi.org/project/pytest-qt/) + [pytest-cov 7.1.0](https://pypi.org/project/pytest-cov/) -- test framework
- [pre-commit 4.6.0 on PyPI](https://pypi.org/project/pre-commit/) -- git hooks
- [PyInstaller 6.20.0 docs](https://pyinstaller.org/en/stable/) -- bundling
- [PyInstaller issue #6387](https://github.com/pyinstaller/pyinstaller/issues/6387) + [#3890](https://github.com/pyinstaller/pyinstaller/issues/3890) -- QtWebEngine pitfalls
- [Inno Setup 6.7.1](https://jrsoftware.org/isdl.php) -- installer
- [Qt QWebEnginePage.acceptNavigationRequest](https://doc.qt.io/qtforpython-6.8/PySide6/QtWebEngineCore/QWebEnginePage.html) -- navigation hardening
- [Qt QWebEngineUrlRequestInterceptor](https://doc.qt.io/qtforpython-6/PySide6/QtWebEngineCore/QWebEngineUrlRequestInterceptor.html) -- request blocking
- [Qt QWebChannel docs](https://doc.qt.io/qt-6/qwebchannel.html) -- bridge architecture
- [Python packaging: single-source version](https://packaging.python.org/en/latest/discussions/single-source-version/) -- version management
- [Microsoft COM apartment threading](https://learn.microsoft.com/en-us/windows/win32/com/multithreaded-apartments) -- COM pitfall documentation
- [PySide6 6.11.0 on PyPI](https://pypi.org/project/PySide6/) -- stack version validation
- Direct codebase analysis of main.py, webview.py, bridge.py, workers.py, settings_state.py, build-installer.ps1, Windows Toolbox.spec -- pitfalls 1-17

### Secondary (MEDIUM confidence)

- [Biome v2 docs](https://biomejs.dev/) -- JS tooling (newer tool, MEDIUM-HIGH)
- [ESLint v10 / eslint-plugin-react issue #3977](https://github.com/jsx-eslint/eslint-plugin-react/issues/3977) -- rationale for Biome over ESLint
- [PySide6 best practices (ZynU)](https://www.zynu.net/ai-skills/pyside6-best-practices) -- MainWindow wiring hub pattern
- [KDAB: Qt WebChannel bridging](https://www.kdab.com/qt-webchannel-bridging-gap-cqml-web/) -- architecture recommendations
- [comtypes gen_py cache corruption issue #182](https://github.com/enthought/comtypes/issues/182) -- COM cache pitfall
- [Cloudscape unsaved changes pattern](https://cloudscape.design/patterns/general/unsaved-changes/) -- draft model UX

---

*Research completed: 2026-04-24*
*Ready for roadmap: yes*