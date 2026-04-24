# Phase 1: Hygiene and Build Pipeline - Context

**Gathered:** 2026-04-24
**Status:** Ready for planning

<domain>
## Phase Boundary

Remove all stale "Windows Toolbox" naming, create standard repository files (.gitignore, README, LICENSE, CLAUDE.md), consolidate version strings to a single source of truth, remove deprecated Qt attributes, and establish a reproducible build pipeline that produces the frontend, PyInstaller app, and Inno Setup installer from a clean checkout.

</domain>

<decisions>
## Implementation Decisions

### Version consolidation
- **D-01:** Version defined once in `app_config.py` as `APP_VERSION`. Frontend receives it via Vite `define` at build time. Inno Setup receives it via `/D` flag from build script. `package.json` version is independent (frontend is private, never published to npm).
- **D-02:** Add all product metadata constants to `app_config.py`: `APP_NAME`, `APP_DISPLAY_NAME`, `APP_VERSION`, `APP_ID`, `APP_EXECUTABLE_NAME`, `APP_DIST_DIR_NAME`, `APP_PUBLISHER`, `APP_SUPPORT_URL`, `APP_SETTINGS_ORG`, `APP_LOG_DIR`, `APP_LOG_FILE`.

### Stale naming
- **D-03:** Rename `Windows Toolbox.spec` to `Virelo.spec`. Update all references in `scripts/build-installer.ps1`.
- **D-04:** Remove all internal migration-phase comments (Phase 7, SC-6, IC-11, v2 labels) from production source files.
- **D-05:** Grep verification: `grep -rn "Windows Toolbox" .` and `grep -rn "Toolbox" .` must return zero matches after cleanup.

### Build script architecture
- **D-06:** Multiple specialized PowerShell scripts in `scripts/`: `bootstrap.ps1`, `clean.ps1`, `build-frontend.ps1`, `build-app.ps1`, `build-installer.ps1`, `verify-release.ps1`. Each script validates preconditions and fails early with clear error messages.
- **D-07:** `build-frontend.ps1` runs `npm ci` if `node_modules` absent, then `npm run build`, then verifies `frontend/dist/index.html` exists.
- **D-08:** `build-app.ps1` calls `build-frontend.ps1` first, then runs PyInstaller with `Virelo.spec`, then verifies `dist/Virelo/Virelo.exe` exists.
- **D-09:** `build-installer.ps1` calls `build-app.ps1` first, locates ISCC.exe, runs `installer/virelo.iss` with version from `app_config.py`, verifies installer output.

### Repository files
- **D-10:** `.gitignore` based on GitHub Python template, extended with: `frontend/node_modules/`, `frontend/dist/`, `build/`, `dist/`, `*.spec.bak`, `.pytest_cache/`, `.ruff_cache/`, `.coverage`, `crash.log`. Do not commit `frontend/dist/` — treat as build artifact.
- **D-11:** `README.md` covers: what Virelo does, Windows-only requirement, admin privilege requirement, personal-use status, build from source instructions, dev mode instructions, planned features section.
- **D-12:** `LICENSE` file with MIT license.
- **D-13:** `CLAUDE.md` provides build commands, project structure, naming conventions, forbidden changes (no stale names, no fake features, no generated artifacts in git), and known footguns.

### Deprecated Qt cleanup
- **D-14:** Remove `setAttribute(AA_EnableHighDpiScaling)` and `setAttribute(AA_UseHighDpiPixmaps)` calls from `main.py`. These are no-ops in Qt 6 and produce deprecation warnings.

### Claude's Discretion
- Build script error message formatting and verbosity level
- README structure and section ordering
- CLAUDE.md organization and level of detail
- Whether `bootstrap.ps1` creates a `.venv` or uses the system Python

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Project vision, constraints, key decisions
- `.planning/REQUIREMENTS.md` — All 52 v1 requirements with phase mapping

### Current codebase
- `.planning/codebase/STACK.md` — Current technology stack analysis
- `.planning/codebase/ARCHITECTURE.md` — Current architecture with layer descriptions
- `.planning/codebase/STRUCTURE.md` — Current file structure
- `.planning/codebase/CONCERNS.md` — Known concerns and issues

### Research
- `.planning/research/STACK.md` — Recommended tooling (ruff, pytest, Biome, pyproject.toml)
- `.planning/research/FEATURES.md` — Table stakes features and dependency ordering
- `.planning/research/PITFALLS.md` — Build pipeline and PyInstaller pitfalls

### Existing build files
- `Windows Toolbox.spec` — Current PyInstaller spec (to be renamed to Virelo.spec)
- `scripts/build-installer.ps1` — Current build script (to be updated)
- `installer/virelo.iss` — Inno Setup script (version to be externalized)
- `frontend/vite.config.js` — Vite config (base: "./" already correct)
- `app_config.py` — Current app config (to be extended with metadata)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `app_config.py` — Already has `APP_NAME`, `ORGANIZATION`, `APP_ID`, `DEFAULTS`. Extend with version and metadata.
- `frontend/vite.config.js` — Already uses `base: "./"` for relative paths. Add `define` for version injection.
- `scripts/build-installer.ps1` — Existing build script to refactor into the multi-script pipeline.

### Established Patterns
- PowerShell for build automation (existing pattern in `scripts/`)
- Inno Setup Pascal Script for installer customization (existing in `installer/virelo.iss`)
- `resource_path()` function in `main.py` for PyInstaller-aware path resolution

### Integration Points
- PyInstaller spec `datas=` list — must include `frontend/dist` and `icon.ico`
- `installer/virelo.iss` `#define MyAppVersion` — currently hardcoded to "1.4.2"
- `frontend/src/pages.jsx` and `frontend/src/panels.jsx` — version strings displayed in About page and elsewhere

</code_context>

<specifics>
## Specific Ideas

- Keep product name "Virelo" — abstract enough for future personal utilities
- Use `pyproject.toml` for project metadata even though full Python packaging isn't needed yet (sets up for Phase 3 quality tooling)
- The user's analysis document specifies exact .gitignore entries and script behaviors — follow those specifications

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 01-hygiene-and-build-pipeline*
*Context gathered: 2026-04-24*
