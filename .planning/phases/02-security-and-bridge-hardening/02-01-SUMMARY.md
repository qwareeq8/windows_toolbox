---
phase: 02-security-and-bridge-hardening
plan: 01
subsystem: webengine-host
tags: [security, webengine, navigation-filter, dev-mode, error-page]
dependency_graph:
  requires: []
  provides: [SEC-01, SEC-02, SEC-03, SEC-04, SEC-05]
  affects: [webview.py]
tech_stack:
  added: []
  patterns: [acceptNavigationRequest-override, context-menu-policy, setHtml-error-page]
key_files:
  created: []
  modified: [webview.py]
decisions:
  - "Dev mode requires explicit VIRELO_DEV=1 -- sys.frozen fallback removed"
  - "data: scheme allowed in navigation filter to support setHtml error pages"
  - "Context menu import deferred inside if-block to avoid top-level import when not needed"
metrics:
  duration: 3min
  completed: "2026-04-24T18:47:32Z"
---

# Phase 02 Plan 01: WebEngine Security Hardening Summary

WebEngine host hardened with navigation filtering, strict dev mode, conditional remote URL access, context menu suppression, and styled error page for missing frontend builds.

## What Changed

### Task 1: Tighten dev mode and add navigation filter (88251e7)
- Removed `sys.frozen` fallback from `_is_dev_mode()` -- now only returns True when `VIRELO_DEV` env var is explicitly set to `1`, `true`, or `yes`
- Added `acceptNavigationRequest` override to `VireloWebPage` that allows `file://` and `data:` unconditionally, `http`/`https` to `localhost` only in dev mode, and blocks all other navigation with logging
- Changed `LocalContentCanAccessRemoteUrls` from unconditional `True` to `_is_dev_mode()` -- disabled in release mode
- Added `setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)` in release mode to prevent Inspect Element access

### Task 2: Missing frontend error page (805f9b2)
- Added `_MISSING_FRONTEND_HTML` module-level constant with a dark-themed error page showing "Frontend build not found" and build instructions
- Modified `_get_frontend_url()` to return `None` when `frontend/dist/index.html` does not exist in release mode
- Updated `VireloWebView.__init__()` to call `page.setHtml(_MISSING_FRONTEND_HTML)` when URL is None
- Updated `reload_frontend()` to handle the None case with `self.page().setHtml()`

## Security Requirements Addressed

| Requirement | Implementation |
|-------------|---------------|
| SEC-01: Block external navigation | `acceptNavigationRequest` blocks all non-local URLs |
| SEC-02: Disable remote URL access in release | `LocalContentCanAccessRemoteUrls` set to `_is_dev_mode()` |
| SEC-03: Strict dev mode detection | `_is_dev_mode()` checks only `VIRELO_DEV` env var |
| SEC-04: Missing frontend error page | `_MISSING_FRONTEND_HTML` rendered via `setHtml()` |
| SEC-05: Context menu disabled in release | `Qt.ContextMenuPolicy.NoContextMenu` in release mode |

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Updated stale comment on dev mode detection**
- **Found during:** Task 1
- **Issue:** Comment on line 29 said "check if VIRELO_DEV env var is set, or if we're not frozen" -- the "or if we're not frozen" part was stale after removing the sys.frozen fallback
- **Fix:** Updated comment to "requires explicit VIRELO_DEV=1 environment variable"
- **Files modified:** webview.py
- **Commit:** 88251e7

## Commits

| Task | Commit | Message |
|------|--------|---------|
| 1 | 88251e7 | feat(02-01): harden WebEngine with navigation filter, strict dev mode, and context menu policy |
| 2 | 805f9b2 | feat(02-01): add missing frontend error page with styled HTML fallback |

## Self-Check: PASSED
