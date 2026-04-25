# Phase 8: Window Chrome and Release - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 08-window-chrome-and-release
**Areas discussed:** Title bar drag zones, Smoke test architecture, Release verification scope, Documentation structure
**Mode:** --auto (all areas auto-selected, recommended options auto-chosen)

---

## Title Bar Drag Zones

| Option | Description | Selected |
|--------|-------------|----------|
| Extend WM_NCHITTEST | Add HTCAPTION to existing nativeEvent handler for top ~40px | ✓ |
| Separate event filter | Install Qt event filter for drag detection | |
| CSS -webkit-app-region | Frontend CSS coordination with Python | |

**User's choice:** [auto] Extend WM_NCHITTEST (recommended default)
**Notes:** Simplest approach — keeps all hit-test logic in one handler, reuses existing pattern, avoids cross-layer coordination complexity.

---

## Smoke Test Architecture

| Option | Description | Selected |
|--------|-------------|----------|
| Full QApp boot, headless | Launch QApplication, construct MainWindow, verify subsystems, exit before event loop | ✓ |
| Python-only checks | Lightweight import/file checks without Qt | |
| Full boot with window | Visible window for manual inspection | |

**User's choice:** [auto] Full QApp boot, headless (recommended default)
**Notes:** Verifies the actual initialization path including Qt, QWebEngine, Settings, and Bridge — catches real failures that import-only checks would miss.

---

## Release Verification Scope

| Option | Description | Selected |
|--------|-------------|----------|
| Expand existing script | Add checks to scripts/verify-release.ps1 | ✓ |
| New Python script | pytest-style verification | |
| Separate pre/post scripts | Split by build stage | |

**User's choice:** [auto] Expand existing script (recommended default)
**Notes:** Script already exists with the right structure. Adding version cross-checks and bundled content checks keeps the pipeline consistent.

---

## Documentation Structure

| Option | Description | Selected |
|--------|-------------|----------|
| README + docs/ | README for overview/quickstart, docs/ for BUILD.md, TROUBLESHOOTING.md, RELEASE.md | ✓ |
| README only | All content in one file | |
| GitHub Wiki | External wiki pages | |

**User's choice:** [auto] README + docs/ (recommended default, matches CHRM-06 requirement)
**Notes:** CHRM-06 explicitly requires "public docs exist in README.md and docs/" — this structure satisfies the requirement while keeping focused, discoverable files.

---

## Claude's Discretion

- Exact pixel height for title bar drag constant
- Window control button exclusion method
- Smoke test output formatting
- README section ordering and detail level
- docs/ file content depth

## Deferred Ideas

None — discussion stayed within phase scope.
