# Phase 5: Bridge and Settings Correctness - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 05-bridge-and-settings-correctness
**Mode:** auto
**Areas discussed:** Dirty state signal, Key capture draft, Launch at login flow, Theme coherence, Boolean strictness, UI preference persistence

---

## Dirty State Signal (BRDG-01)

| Option | Description | Selected |
|--------|-------------|----------|
| Dedicated `dirty_changed(bool)` signal | Separate signal on VireloBridge, emits on draft state transitions | ✓ |
| Extend `settings_changed` payload | Add `has_draft` field to the existing settings JSON payload | |

**User's choice:** [auto] Dedicated `dirty_changed(bool)` signal (recommended default)
**Notes:** Cleaner separation of concerns. Dirty state is orthogonal to settings values. Matches existing Qt signal pattern.

---

## Key Capture Draft Integration (BRDG-03)

| Option | Description | Selected |
|--------|-------------|----------|
| Route through `apply_draft()` | Captured key goes to draft, shown as pending, committed on save | ✓ |
| Direct write (bypass draft) | Keep current behavior, key takes effect immediately | |

**User's choice:** [auto] Route through `apply_draft()` (recommended default — matches BRDG-03 requirement text)
**Notes:** BRDG-03 explicitly says "reflected as a pending (unsaved) draft change." HotkeyListener update deferred to commit for consistency.

---

## Launch at Login Flow (BRDG-02)

| Option | Description | Selected |
|--------|-------------|----------|
| Through draft/commit with side effects | Toggle goes to draft, shortcut created/removed on commit | ✓ |
| Immediate shortcut toggle | Keep current behavior, toggle creates/removes shortcut immediately | |

**User's choice:** [auto] Through draft/commit with side effects (recommended default)
**Notes:** Consistent with all other settings. Error reporting via snap_status signal.

---

## Theme Coherence (BRDG-05)

| Option | Description | Selected |
|--------|-------------|----------|
| Return both mode + effective in `get_theme_mode` | JSON `{"mode": "system", "effective": "dark"}` + keep `theme_applied` signal | ✓ |
| Two separate signals | New `theme_mode_changed(str)` alongside `theme_applied(str)` | |
| Combined JSON signal | Single signal with both mode and effective | |

**User's choice:** [auto] Return both mode + effective in `get_theme_mode`, route through draft (recommended default)
**Notes:** Simplest approach — existing `theme_applied` for rendering, settings payload for mode. Removes separate `apply_theme` slot.

---

## Boolean Strictness (BRDG-04)

| Option | Description | Selected |
|--------|-------------|----------|
| `_strict_bool()` at bridge boundary | Accepts only True/False/"true"/"false"/1/0, raises ValueError otherwise | ✓ |
| Tighten `_safe_bool` to raise on ambiguous | Make `_safe_bool` reject "yes"/"no"/"on"/"off" | |

**User's choice:** [auto] `_strict_bool()` at bridge boundary (recommended default)
**Notes:** Keep `_safe_bool` permissive for QSettings reads (Registry stores weird types). Strict at bridge boundary where frontend values enter.

---

## UI Preference Persistence (BRDG-06)

| Option | Description | Selected |
|--------|-------------|----------|
| Add accent/density/minimize-to-tray to Python model | Persist visible controls, leave radius/sidebarMode as frontend constants | ✓ |
| Remove accent/density controls | Strip the controls from GeneralPage | |

**User's choice:** [auto] Add to Python model (recommended default)
**Notes:** Users expect accent and density to survive restarts. BRDG-06 requirement says "persisted OR removed" — persisted is the better UX.

---

## Claude's Discretion

- Debounce on `dirty_changed`, accent/density validation approach, side-effect refactoring, tray sync timing, theme revert animation

## Deferred Ideas

None — auto mode stayed within phase scope

## Auto-Resolved

- All 6 gray areas auto-resolved with recommended defaults (no Unclear assumptions)
