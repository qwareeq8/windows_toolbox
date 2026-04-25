# Phase 7: CI and Repo Hygiene - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 07-ci-and-repo-hygiene
**Areas discussed:** CI platform strategy, Stale-name false positive fix, Version sync mechanism, Class rename scope
**Mode:** --auto (all decisions auto-selected)

---

## CI Platform Strategy

| Option | Description | Selected |
|--------|-------------|----------|
| Platform skip markers | Add pytest.mark.skipif for Win32 tests, keep ubuntu-latest | ✓ |
| Switch to windows-latest | Run all tests on Windows runner | |
| Dual platform matrix | Run on both ubuntu and windows | |

**User's choice:** [auto] Platform skip markers (recommended default)
**Notes:** test_snap_geometry.py::test_restore_maximized_window imports win32con/win32gui/ctypes at lines 110-114 — will fail on Ubuntu. Skip marker is lightest-touch fix. All other unit tests use pure Python or mock PySide6.

---

## Stale-Name False Positive Fix

| Option | Description | Selected |
|--------|-------------|----------|
| Exclude .github/ + rephrase docstring | Targeted exclusion, no detection loss | ✓ |
| Exclude tests/ entirely | Broader exclusion, could miss real stale refs | |
| Restructure grep as Python script | More maintainable, heavier change | |

**User's choice:** [auto] Exclude .github/ + rephrase docstring (recommended default)
**Notes:** ci.yml lines 53/55/62 reference "Windows Toolbox" in the check itself. test_app_config.py line 49 uses it in a docstring. Both are legitimate references, not stale naming.

---

## Version Sync Mechanism

| Option | Description | Selected |
|--------|-------------|----------|
| CI check + build sync | Fail CI on mismatch, sync at build time | ✓ |
| Build-time only sync | Script reads config.py, patches package.json | |
| Manual sync | Developer responsibility | |

**User's choice:** [auto] CI check + build sync (recommended default)
**Notes:** package.json is 1.4.2, config.py is 1.5.0 — already drifted. CI check prevents future drift. Build script can inject version into Vite define.

---

## Class Rename Scope

| Option | Description | Selected |
|--------|-------------|----------|
| Code + tests + CLAUDE.md | All live references consistent | ✓ |
| Code + tests only | Skip doc updates | |
| Code only | Minimal change | |

**User's choice:** [auto] Code + tests + CLAUDE.md (recommended default)
**Notes:** CLAUDE.md is a live project reference used by AI agents — stale class names there would cause confusion. .planning/ docs are historical and should NOT be updated.

---

## Claude's Discretion

- Exact pytest marker placement and import guard structure
- Clean script implementation details
- CI version-check implementation approach
- Commit ordering within phase

## Deferred Ideas

None — discussion stayed within phase scope.
