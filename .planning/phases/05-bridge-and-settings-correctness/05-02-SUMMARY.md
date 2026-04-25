---
phase: 05-bridge-and-settings-correctness
plan: 02
subsystem: bridge, window
tags: [draft-commit, side-effects, key-capture, theme, startup, tray-menu]

# Dependency graph
requires:
  - phase: 05-bridge-and-settings-correctness
    plan: 01
    provides: dirty_changed signal, _strict_bool, minimize_to_tray key
provides:
  - All settings changes route through draft/commit model
  - _apply_side_effects handles run_at_startup, minimize_to_tray, theme
  - Key capture produces draft change (not immediate persist)
  - Theme previews on draft, reverts on discard
affects: [05-03-PLAN, frontend theme/dirty state handling]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Tray menu actions route through apply_draft->commit_draft->_apply_side_effects"
    - "Theme immediate-apply in save_settings, revert in discard_draft"
    - "Lazy import for startup shortcut functions to avoid circular imports"

key-files:
  created: []
  modified:
    - virelo/bridge/bridge.py
    - virelo/app/window.py

key-decisions:
  - "Removed apply_theme slot -- theme changes go through save_settings + apply_draft"
  - "Removed toggle_run_at_startup slot -- startup changes go through tray menu draft/commit"
  - "get_theme_mode now returns {mode, effective} for frontend to distinguish user choice from resolved theme"
  - "Key capture uses apply_draft without commit -- captured key shows as pending unsaved change"
  - "_toggle_theme (Ctrl+T) now only previews visually without persisting (consistent with draft model)"

patterns-established:
  - "Tray menu draft/commit pattern: apply_draft -> commit_draft -> _apply_side_effects in one action"
  - "Lazy import inside _apply_side_effects to break circular bridge<->window dependency"
  - "Theme preview/revert: save_settings applies immediately, discard_draft reverts to persisted"

requirements-completed: [BRDG-02, BRDG-03, BRDG-05]

# Metrics
duration: 5min
completed: 2026-04-25
---

# Phase 5 Plan 02: Draft/Commit Rerouting Summary

**Removed standalone bridge slots, routed key capture/startup/theme through draft/commit model, expanded side effects**

## Performance

- **Duration:** 5 min
- **Started:** 2026-04-25
- **Completed:** 2026-04-25
- **Tasks:** 2
- **Files modified:** 2

## Accomplishments
- Removed `apply_theme` slot from VireloBridge -- theme selection now goes through `save_settings` -> `apply_draft`
- Removed `toggle_run_at_startup` slot from VireloBridge -- startup toggle now routes through tray menu draft/commit
- Updated `get_theme_mode` to return `{mode, effective}` structure for frontend disambiguation
- Added theme immediate-apply in `save_settings` and revert-to-persisted in `discard_draft`
- Expanded `_apply_side_effects` with run_at_startup (shortcut create/remove with error handling), minimize_to_tray (flag update), and tray menu checkbox sync
- Rerouted `_on_capture_key` through `apply_draft` -- captured key is a pending draft change, hotkey listener updates on commit
- Removed `on_key_captured` slot, `key_captured` signal, and its connection
- Updated `_toggle_run_at_startup` and `_toggle_minimize_on_exit` to route through draft/commit/side-effects
- Initialized `minimize_to_tray_on_exit` from persisted settings instead of hardcoded `True`
- Removed direct `self.settings.theme =` write from `_apply_theme_mode`

## Task Commits

1. **Tasks 1+2: Draft/commit rerouting for bridge and window** - `9c370d4` (feat)

## Files Modified
- `virelo/bridge/bridge.py` - Removed apply_theme/toggle_run_at_startup slots, updated get_theme_mode return structure, added theme preview/revert, expanded _apply_side_effects
- `virelo/app/window.py` - Rerouted _on_capture_key through draft, removed on_key_captured/key_captured, updated tray menu handlers, fixed _apply_theme_mode, initialized minimize_to_tray from settings

## Deviations from Plan
None

## Issues Encountered
None -- no venv present so PySide6-dependent tests could not be run, but AST parsing and grep checks all pass.

## Next Phase Readiness
- All settings changes now route through draft/commit model
- Plan 03 (frontend wiring) can rely on dirty_changed signal, structured get_theme_mode response, and no standalone slots

## Self-Check: PASSED

Both modified files verified. Commit 9c370d4 found in git log. All acceptance criteria greps pass.

---
*Phase: 05-bridge-and-settings-correctness*
*Completed: 2026-04-25*
