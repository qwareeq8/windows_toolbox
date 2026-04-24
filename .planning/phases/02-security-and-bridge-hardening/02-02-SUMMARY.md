---
phase: 02-security-and-bridge-hardening
plan: 02
subsystem: bridge, settings
tags: [draft-model, structured-payloads, window-commands, state-management]
dependency_graph:
  requires: []
  provides: [draft-state-model, structured-bridge-payloads, window-command-slot]
  affects: [frontend-save-flow, frontend-settings-parsing]
tech_stack:
  added: []
  patterns: [draft-commit-discard, structured-json-response, command-whitelist]
key_files:
  created: []
  modified:
    - settings_state.py
    - bridge.py
decisions:
  - "apply_partial renamed to apply_draft -- stores in draft dict, not QSettings"
  - "Side effects fire only on commit_draft, not on individual save_settings calls"
  - "Single setWindowCommand slot with string command (not separate minimize/close slots)"
  - "get_launch_at_login and get_snap_enabled changed from result=bool to result=str for structured payloads"
metrics:
  duration: 3min
  completed: 2026-04-24
---

# Phase 2 Plan 2: Bridge Draft State and Response Standardization Summary

Python-side draft/commit state model in SettingsState with structured JSON payloads on all bridge slots and setWindowCommand for title bar controls.

## What Was Done

### Task 1: Add draft state model to SettingsState (3a2e374)
- Added `_draft = None` initialization in `__init__` for staging unsaved changes
- Replaced `apply_partial` with `apply_draft` that validates input and stores in `_draft` dict without calling `Settings.save()`
- Added `commit_draft()` that writes `_draft` values to Settings, calls `save()`, clears draft, returns applied changes
- Added `discard_draft()` that sets `_draft = None` without persisting
- Added `has_draft` property returning whether unsaved changes exist
- Modified `get_all()` to overlay `_draft` onto persisted settings before normalization
- Modified `reset_to_defaults()` to clear `_draft` before resetting
- Unknown keys now produce structured error `{"ok": false, "error": "Unknown keys: [...]"}` instead of being silently dropped

### Task 2: Add bridge slots and standardize responses (1890d89)
- `save_settings` now calls `apply_draft` instead of `apply_partial`, and no longer calls `_apply_side_effects`
- Added `commit_draft` slot: persists draft via `SettingsState.commit_draft()`, emits `settings_changed`, applies side effects
- Added `discard_draft` slot: clears draft via `SettingsState.discard_draft()`, pushes persisted settings to frontend
- Added `has_draft` slot: returns `{"ok": true, "data": bool}`
- Added `setWindowCommand` slot: validates command string against `("minimize", "close")` whitelist, executes on MainWindow
- Standardized `get_settings` from raw JSON to `{"ok": true, "data": {settings}}`
- Standardized `get_theme_mode` from bare string to `{"ok": true, "data": "mode"}`
- Standardized `get_launch_at_login` from bare bool to `{"ok": true, "data": bool}`
- Standardized `get_snap_enabled` from bare bool to `{"ok": true, "data": bool}`
- Standardized `reset_defaults` to `{"ok": true, "data": {settings}}`
- Standardized `toggle_run_at_startup` response with `data` wrapper
- `apply_theme` uses `apply_draft` instead of `apply_partial`

## Deviations from Plan

None -- plan executed exactly as written.

## Decisions Made

1. **apply_partial fully replaced**: The old `apply_partial` method is completely removed. No references remain in either `settings_state.py` or `bridge.py`.

2. **Single setWindowCommand slot**: Chose single slot with string parameter over separate `minimizeWindow`/`closeWindow` slots (per Claude's discretion in CONTEXT.md). More extensible if future commands are added.

3. **Return type changes**: `get_launch_at_login` and `get_snap_enabled` changed from `@Slot(result=bool)` returning bare bools to `@Slot(result=str)` returning structured JSON. This is a breaking change for any frontend code that directly used the bool return.

4. **Side effects deferred to commit**: `_apply_side_effects` is no longer called in `save_settings`. It fires only during `commit_draft`, preventing startup shortcut creation/removal, snap manager updates, and theme application until the user explicitly saves.

## Downstream Impact

The following frontend changes are required in Plan 02-03:
- `bridge.get_settings()` callback must unwrap `result.data` instead of using raw dict
- `bridge.get_theme_mode()` callback must unwrap `result.data` instead of using bare string
- `bridge.get_launch_at_login()` callback must parse JSON and unwrap `result.data`
- `bridge.get_snap_enabled()` callback must parse JSON and unwrap `result.data`
- `handleSave` must call `bridge.commit_draft()` instead of current flow
- `handleDiscard` must call `bridge.discard_draft()`
- Mock bridge in `bridge.js` must add `commit_draft`, `discard_draft`, `has_draft`, `setWindowCommand` methods
- Title bar buttons must call `bridge.setWindowCommand("minimize")` and `bridge.setWindowCommand("close")`

## Requirements Coverage

| Requirement | Status | Implementation |
|-------------|--------|----------------|
| BRDG-01 | Done | `_draft` dict in SettingsState holds unsaved changes separately |
| BRDG-02 | Done | `commit_draft` persists draft and applies side effects |
| BRDG-03 | Done | `discard_draft` clears draft, frontend receives persisted values |
| BRDG-04 | Done | Unknown keys rejected with `{"ok": false, "error": "Unknown keys: [...]"}` |
| BRDG-05 | Done | All slots return `{"ok": true/false, ...}` structured payloads |
| BRDG-06 | Done | Startup shortcut creation/removal fires during `commit_draft` side effects |

## Verification Results

- `apply_partial` removed from both files (grep returns zero matches)
- `commit_draft` present in both `bridge.py` and `settings_state.py`
- `discard_draft` present in both files
- `setWindowCommand` present in `bridge.py`
- Structured `"ok"` count in bridge.py: 37 occurrences (all slots covered)
- `apply_draft` does NOT call `self._settings.save()`
- `commit_draft` DOES call `self._settings.save()`

## Self-Check: PASSED
