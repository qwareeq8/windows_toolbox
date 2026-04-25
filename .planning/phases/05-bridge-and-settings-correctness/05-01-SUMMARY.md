---
phase: 05-bridge-and-settings-correctness
plan: 01
subsystem: settings
tags: [pyside6, qsettings, validation, boolean-parsing, draft-commit, signals]

# Dependency graph
requires:
  - phase: 04-snap-and-explorer-hardening
    provides: draft/commit model in SettingsState, bridge signal-slot patterns
provides:
  - _strict_bool function for bridge-boundary boolean validation
  - accent, density, minimize_to_tray settings keys in DEFAULTS, KEYS, Settings, SettingsState
  - _VALID_ACCENTS and _VALID_DENSITIES validation tuples
  - dirty_changed Signal(bool) on VireloBridge
affects: [05-02-PLAN, 05-03-PLAN, frontend dirty state subscription, bridge slot cleanup]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "_strict_bool coercer replacing Python bool() for bridge-boundary boolean validation"
    - "Allowlist validation with fallback-to-default for enum-like string settings"
    - "dirty_changed Signal(bool) emitted alongside settings_changed for orthogonal dirty tracking"

key-files:
  created: []
  modified:
    - virelo/app/config.py
    - virelo/settings/state.py
    - virelo/settings/persistence.py
    - virelo/bridge/bridge.py
    - tests/unit/test_settings_state.py
    - tests/unit/test_app_config.py

key-decisions:
  - "_strict_bool accepts only True/False/'true'/'false'/1/0 -- rejects 'yes'/'no'/None/'' per D-15"
  - "Accent/density validated via simple tuple membership check, not enum class (matches project style)"
  - "dirty_changed emits on every apply_draft without debounce (D-01 discretion: lightweight signal)"
  - "_safe_bool in persistence.py unchanged per D-17 -- QSettings reads stay permissive"

patterns-established:
  - "_strict_bool: strict bridge-boundary boolean coercion preventing bool('false')==True"
  - "Allowlist validation pattern: _VALID_ACCENTS/_VALID_DENSITIES with DEFAULTS fallback"
  - "Signal(bool) dirty_changed: emits has_draft in save_settings, False in commit/discard/reset"

requirements-completed: [BRDG-01, BRDG-04, BRDG-06]

# Metrics
duration: 4min
completed: 2026-04-25
---

# Phase 5 Plan 01: Python Settings Model Foundation Summary

**Strict boolean parsing via _strict_bool, three new settings keys (accent/density/minimize_to_tray) with validation, and dirty_changed Signal(bool) on VireloBridge**

## Performance

- **Duration:** 4 min
- **Started:** 2026-04-25T00:21:25Z
- **Completed:** 2026-04-25T00:25:35Z
- **Tasks:** 2
- **Files modified:** 6

## Accomplishments
- Added _strict_bool function preventing the bool("false")==True bug at the bridge boundary, replacing bare bool in all 5 boolean KEYS entries
- Extended settings model with accent (allowlist: slate/teal/blue/rust/purple), density (allowlist: compact/cozy/comfortable), and minimize_to_tray (strict bool) across DEFAULTS, KEYS, Settings QSettings read/write, and SettingsState
- Added dirty_changed = Signal(bool) on VireloBridge, emitted in all 4 draft-modifying methods (save_settings, commit_draft, discard_draft, reset_defaults)
- Added 13 new unit tests covering _strict_bool behavior, accent/density validation, minimize_to_tray strict bool, and commit persistence of new keys -- all 66 tests pass

## Task Commits

Each task was committed atomically:

1. **Task 1 (TDD RED): Failing tests for _strict_bool and new keys** - `abce4af` (test)
2. **Task 1 (TDD GREEN): Implement _strict_bool, new KEYS, accent/density validation** - `11e13ef` (feat)
3. **Task 2: Add dirty_changed signal and wire emission** - `3644ea7` (feat)

## Files Created/Modified
- `virelo/app/config.py` - Added accent, density, minimize_to_tray to DEFAULTS dict
- `virelo/settings/state.py` - Added _strict_bool function, _VALID_ACCENTS/_VALID_DENSITIES tuples, replaced bool with _strict_bool in KEYS, added new key entries, added accent/density validation in apply_draft
- `virelo/settings/persistence.py` - Added QSettings read/write for accent, density, minimize_to_tray
- `virelo/bridge/bridge.py` - Added dirty_changed = Signal(bool), emit calls in save_settings/commit_draft/discard_draft/reset_defaults
- `tests/unit/test_settings_state.py` - Added 13 new tests for _strict_bool and new settings keys
- `tests/unit/test_app_config.py` - Updated test_defaults_has_all_keys to expect 14 keys

## Decisions Made
- Used _strict_bool as a standalone function rather than a class method, matching the module-level normalize_theme_mode/normalize_snap_presses pattern
- Accent/density validated via simple tuple membership with fallback to DEFAULTS, not enum class (project convention)
- dirty_changed emits on every apply_draft without debounce -- the signal is a single boolean and React re-render cost is negligible

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Updated test_defaults_has_all_keys to expect 14 keys**
- **Found during:** Task 1 (GREEN phase, full test suite run)
- **Issue:** Existing test_app_config.py hardcoded expected_keys set with 11 entries; adding 3 new DEFAULTS entries caused assertion failure
- **Fix:** Added accent, density, minimize_to_tray to the expected_keys set and updated docstring count
- **Files modified:** tests/unit/test_app_config.py
- **Verification:** All 66 unit tests pass
- **Committed in:** 11e13ef (Task 1 GREEN commit)

---

**Total deviations:** 1 auto-fixed (1 bug)
**Impact on plan:** Necessary test update to match expanded DEFAULTS. No scope creep.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- _strict_bool and new KEYS are ready for Plan 02 (bridge slot cleanup, key capture via draft)
- dirty_changed signal is declared and emitted; Plan 03 will wire the frontend subscription
- All boolean KEYS now use _strict_bool, so Plans 02/03 can rely on strict boolean validation at the bridge boundary

## Self-Check: PASSED

All 6 files verified present. All 3 commit hashes (abce4af, 11e13ef, 3644ea7) found in git log.

---
*Phase: 05-bridge-and-settings-correctness*
*Completed: 2026-04-25*
