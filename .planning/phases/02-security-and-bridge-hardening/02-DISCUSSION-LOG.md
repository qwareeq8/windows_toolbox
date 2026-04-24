# Phase 2: Security and Bridge Hardening - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 02-security-and-bridge-hardening
**Areas discussed:** WebEngine lockdown, Draft state architecture, Fake control removal, Bridge response standardization, Title bar and command palette wiring
**Mode:** Auto (all decisions auto-selected with recommended defaults)

---

## WebEngine Lockdown

| Option | Description | Selected |
|--------|-------------|----------|
| acceptNavigationRequest override | Block non-local URLs in VireloWebPage subclass | ✓ |
| URL whitelist in settings | Configurable allow-list for external URLs | |
| No blocking (current) | Keep existing behavior with no restrictions | |

**User's choice:** [auto] acceptNavigationRequest override — blocks all non-local navigation
**Notes:** Combined with conditional LocalContentCanAccessRemoteUrls (False in release, True in dev) and dev mode requiring explicit VIRELO_DEV=1 env var.

| Option | Description | Selected |
|--------|-------------|----------|
| Styled inline HTML error | Show build instructions when frontend/dist missing | ✓ |
| Blank page with console error | Log error, show nothing visible | |
| MessageBox before load | Qt dialog warning before attempting load | |

**User's choice:** [auto] Styled inline HTML error page — developer-friendly, explains how to build
**Notes:** Context menu disabled in release, enabled in dev for Inspect Element access.

---

## Draft State Architecture

| Option | Description | Selected |
|--------|-------------|----------|
| Python-side draft dict | SettingsState holds _draft overlay, commit/discard slots | ✓ |
| Frontend-only draft | React state is the draft, Python saves on commit | |
| Undo stack | Full undo/redo with history | |

**User's choice:** [auto] Python-side draft dict — changes stored in _draft, separate from persisted Settings
**Notes:** Commit writes to QSettings and applies side effects. Discard reverts to persisted values. get_settings returns merged view (persisted + draft overlay).

---

## Fake Control Removal

| Option | Description | Selected |
|--------|-------------|----------|
| Remove entirely | Delete all fake controls, no stubs | ✓ |
| Disable with "coming soon" | Grey out controls with tooltip | |
| Implement backends | Build real backend support for each | |

**User's choice:** [auto] Remove entirely — per CLAUDE.md forbidden changes rule
**Notes:** Removes: rememberCols, showHidden, showExts, startTray, autoUpdate, telemetry toggles. Removes: "Check for updates" button, "Up to date" badge, "Documentation Open", "Report an issue Open" buttons. Removes hardcoded changelog. Explorer page keeps only auto-size toggle. Shortcuts page keeps only real shortcuts (3 items).

---

## Bridge Response Standardization

| Option | Description | Selected |
|--------|-------------|----------|
| Uniform {ok, data/error} | All slots return consistent structure | ✓ |
| Status codes | HTTP-style numeric codes + message | |
| Keep mixed (current) | Some return {error}, some {ok, error} | |

**User's choice:** [auto] Uniform {ok, data/error} — consistent across all bridge slots
**Notes:** Unknown keys in apply_partial rejected with structured error (not silently dropped). All inputs validated before acting.

---

## Title Bar and Command Palette Wiring

| Option | Description | Selected |
|--------|-------------|----------|
| Bridge slots for window ops | setWindowCommand bridge slot, palette wired to handlers | ✓ |
| Direct Qt calls from JS | JavaScript calls window.close() etc | |
| Keep inert (current) | Title bar buttons remain non-functional | |

**User's choice:** [auto] Bridge slots — minimize/close via setWindowCommand, palette actions wired to existing save/reset/test-snap handlers
**Notes:** Maximize button removed (frameless window, no maximize behavior). Command palette "Test snap" -> bridge.test_snap(), "Reset to defaults" -> handleReset, "Save changes" -> handleSave.

---

## Claude's Discretion

- Error page HTML styling and copy
- Internal draft implementation (dict overlay vs full copy)
- Whether to use single setWindowCommand or separate minimize/close slots
- Sidebar status text after fake update badge removal

## Deferred Ideas

None — discussion stayed within phase scope
