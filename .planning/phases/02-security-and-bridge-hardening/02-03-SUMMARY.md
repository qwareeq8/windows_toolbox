---
phase: 02-security-and-bridge-hardening
plan: 03
subsystem: frontend
tags: [fake-control-removal, draft-model-wiring, command-palette, title-bar, structured-responses]
dependency_graph:
  requires: [draft-state-model, structured-bridge-payloads, window-command-slot]
  provides: [clean-frontend-pages, draft-aware-save-flow, wired-command-palette, functional-title-bar]
  affects: [frontend-user-experience, bridge-response-handling]
tech_stack:
  added: []
  patterns: [structured-response-unwrapping, draft-commit-discard-ui-flow, optional-chaining-callbacks]
key_files:
  created: []
  modified:
    - frontend/src/pages.jsx
    - frontend/src/app.jsx
    - frontend/src/panels.jsx
    - frontend/src/bridge.js
    - frontend/src/main.jsx
decisions:
  - "set() sends partial state to Python draft on every change via bridge.save_settings"
  - "TitleBar close button hover uses Windows-standard red (#e81123)"
  - "Command palette action handlers use optional chaining (onSave?.()) for safety"
  - "useMemo dependency array updated to include onTestSnap, onSave, onReset for correctness"
metrics:
  duration: 5min
  completed: 2026-04-24
---

# Phase 2 Plan 3: Frontend Cleanup and Draft Model Wiring Summary

Removed all fake UI controls, wired Save/Discard to commit_draft/discard_draft, connected title bar and command palette to real bridge methods, and updated all bridge response handling to unwrap structured {ok, data} payloads.

## What Was Done

### Task 1: Remove fake controls and clean up pages (bd51a4f)
- **ExplorerPage**: Removed "Remember column widths" toggle and entire "Hidden files" card (show hidden files + show file extensions). Now contains only "Auto-size columns on folder change" toggle.
- **ShortcutsPage**: Removed 5 fake shortcuts ("Cycle preset sizes", "Move to next display", "Move to previous display", "Toggle Virelo pause", "Open settings"). Kept 3 real items: "Trigger snap", "Restore last snap", "Command palette".
- **GeneralPage**: Removed "Start minimized to tray", "Automatic updates", "Anonymous telemetry" rows. Renamed "Startup & updates" card to "Startup". Each remaining card has a single row marked `last`.
- **AboutPage**: Removed changelog card, "Check for updates" button, "Up to date" badge, "Documentation" row, "Report an issue" row. Now shows only icon, name, version, and License row.
- **bridgeToState**: Removed 6 fake keys: `rememberCols`, `showHidden`, `showExts`, `startTray`, `autoUpdate`, `telemetry`.
- **stateToBridge**: Removed `theme: undefined` line.
- **Initial useState**: Removed same 6 fake keys from initial state object.
- **Sidebar**: Simplified version text from `v{version} . up to date` (with green dot) to just `v{version}`.
- Removed unused `Badge` import from pages.jsx.

### Task 2: Wire draft model, title bar, command palette, and update mock bridge (6641fa3)
- **handleSave**: Now calls `bridge.commit_draft()` instead of `bridge.save_settings(stateToBridge(state))`.
- **handleDiscard**: Now calls `bridge.discard_draft()` instead of `bridge.get_settings()`. The `settings_changed` signal handler (unchanged) updates React state when discard pushes persisted values.
- **handleReset**: Unwraps structured response -- uses `r.data` from `bridge.reset_defaults` response.
- **Initial get_settings**: Unwraps structured response -- uses `r.data` from initial load.
- **set() function**: Now sends partial state updates to Python draft via `bridge.save_settings(stateToBridge(next))` on every change, keeping the draft in sync.
- **TitleBar**: Replaced static `['--','[]','x']` divs with two functional buttons (minimize and close) calling `bridge.setWindowCommand()`. Maximize button removed. Close button hover uses Windows-standard red. TitleBar accepts `bridge` as a prop.
- **CommandPalette**: "Test snap" calls `onTestSnap()`, "Save changes" calls `onSave()`, "Reset to defaults" calls `onReset()`. No more no-ops or fake `_saved` state. `useMemo` deps updated.
- **bridge.js MOCK_BRIDGE**: Added `commit_draft`, `discard_draft`, `has_draft`, `setWindowCommand` slots. All existing slots updated to return structured `{ok, data}` payloads (get_settings wraps MOCK_SETTINGS, get_theme_mode returns `{ok: true, data: 'dark'}`, etc.).
- **main.jsx**: Updated `get_theme_mode` callback to parse and unwrap structured response.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] Fixed get_theme_mode structured response in main.jsx**
- **Found during:** Task 2
- **Issue:** `main.jsx` called `bridge.get_theme_mode()` and used the raw callback value as a string (`mode === 'light'`). After Plan 02-02 changed `get_theme_mode` to return `{"ok": true, "data": "dark"}`, the raw value would be a JSON string, not a bare theme mode.
- **Fix:** Updated `main.jsx` to JSON.parse the result and unwrap `r.data` before comparing theme modes. Added fallback to 'dark' on parse failure.
- **Files modified:** `frontend/src/main.jsx`
- **Commit:** 6641fa3

**2. [Rule 1 - Bug] Wired set() to send changes to Python draft**
- **Found during:** Task 2
- **Issue:** The `set()` function only updated local React state without sending changes to the Python draft via `bridge.save_settings()`. This meant `commit_draft` would have no draft to persist.
- **Fix:** Updated `set()` to call `bridge.save_settings(stateToBridge(next))` on every state change, keeping the Python draft in sync with the UI.
- **Files modified:** `frontend/src/app.jsx`
- **Commit:** 6641fa3

## Decisions Made

1. **set() sends full state to draft on every change**: Rather than computing a diff of changed keys, `set()` sends the full state via `stateToBridge()` to `bridge.save_settings()`. The Python side's `apply_draft` handles the merge. This is simpler and avoids bugs from partial key mapping.

2. **Close button hover color**: Used `#e81123` (Windows-standard close button red) for hover state, matching native Windows title bar behavior.

3. **Optional chaining for command palette callbacks**: Used `onTestSnap?.()` pattern to safely handle cases where props might not be provided.

4. **useMemo deps updated**: Added `onTestSnap`, `onSave`, `onReset` to the `useMemo` dependency array for React correctness.

## Requirements Coverage

| Requirement | Status | Implementation |
|-------------|--------|----------------|
| UI-01 | Done | All fake controls removed from ExplorerPage, GeneralPage, AboutPage |
| UI-02 | Done | Fake state keys removed from bridgeToState, stateToBridge, initial state |
| UI-03 | Done | About page shows only icon, name, version, license |
| UI-04 | Done | Command palette actions wired to handleTestSnap, handleSave, handleReset |
| UI-05 | Done | Title bar minimize/close call bridge.setWindowCommand; maximize removed |
| UI-06 | Done | Shortcuts page has exactly 3 items: Trigger snap, Restore last snap, Command palette |

## Verification Results

- Zero matches for fake keys in pages.jsx and app.jsx (rememberCols, showHidden, showExts, startTray, autoUpdate, telemetry)
- `commit_draft` present in app.jsx and bridge.js
- `discard_draft` present in app.jsx and bridge.js
- `setWindowCommand` present in app.jsx and bridge.js
- `onTestSnap`, `onSave`, `onReset` present in panels.jsx
- Zero matches for "up to date", "Check for updates", "Anonymous telemetry", "Automatic updates" across frontend/src/
- `r.data` present in app.jsx (structured response unwrapping)

## Self-Check: PASSED
