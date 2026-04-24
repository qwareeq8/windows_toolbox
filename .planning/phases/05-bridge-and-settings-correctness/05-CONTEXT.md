# Phase 5: Bridge and Settings Correctness - Context

**Gathered:** 2026-04-24 (auto mode)
**Status:** Ready for planning

<domain>
## Phase Boundary

Every setting flows through a single Python-owned draft/commit model with correct types and coherent signals. The frontend reflects Python-driven dirty state, captured keys go through draft, theme changes emit both mode and effective theme, booleans parse strictly, and UI preferences (accent, density, minimize-to-tray) are persisted.

</domain>

<decisions>
## Implementation Decisions

### Dirty state signal (BRDG-01)
- **D-01:** Add a `dirty_changed(bool)` signal on `VireloBridge`. Emit `True` when `apply_draft()` creates or extends a draft, `False` when `commit_draft()` or `discard_draft()` clears it. Also emit `False` after `reset_to_defaults()`.
- **D-02:** Frontend replaces its local `unsaved` React state with a subscription to `dirty_changed`. The `setUnsaved(true)` call in `app.jsx:set()` and the `setUnsaved(false)` calls in `handleSave`/`handleDiscard`/`handleReset` are removed — the signal drives the footer dirty indicator.
- **D-03:** The `settings_changed` signal continues to push the full settings dict (including draft overlay). `dirty_changed` is a separate signal because dirty state is orthogonal to settings values.

### Key capture via draft (BRDG-03)
- **D-04:** `_on_capture_key` in `window.py` routes the captured key through `self._settings_state.apply_draft({"snap_key": key_str})` (or `"restore_key"`) instead of writing directly to `self.settings.snap_key`. This triggers `dirty_changed(True)` — the user sees the new key as a pending unsaved change.
- **D-05:** The captured key does NOT immediately update the `HotkeyListener` binding. The old binding stays active until `commit_draft`. Hotkey listener updates happen in `_apply_side_effects` on commit (already handled for `snap_key`/`restore_key`). This prevents the confusing state of a working-but-unsaved key binding.
- **D-06:** On discard, the key reverts to the previously saved value. `discard_draft()` clears `_draft`, and `settings_changed` pushes persisted values — the frontend re-renders with the old key.
- **D-07:** The `on_key_captured` slot in `MainWindow` that directly sets `self.settings.snap_key` and emits `snap_status` is removed. All capture result handling goes through `_on_capture_key` which now uses the draft path.

### Launch at login side effects (BRDG-02)
- **D-08:** The `toggle_run_at_startup` bridge slot is removed. Launch at login changes flow through the normal `save_settings` → `apply_draft` path like all other settings. The General page's "Launch at login" toggle calls `app.set({ launchLogin: v })`.
- **D-09:** `_apply_side_effects` gains a `run_at_startup` handler: when `run_at_startup` is in the committed dict, call `create_startup_shortcut()` or `remove_startup_shortcut()` accordingly. Wrap in try/except and report errors via `snap_status` signal with a descriptive message (e.g., "Failed to create startup shortcut: {error}").
- **D-10:** The tray menu "Run at Startup" action syncs with the persisted setting on startup and after each `commit_draft`. It does NOT directly toggle the shortcut — it goes through the bridge like the frontend control.

### Theme coherence (BRDG-05)
- **D-11:** `get_theme_mode` slot returns JSON `{"mode": "system", "effective": "dark"}` — both the user's chosen mode and the resolved effective theme. The frontend uses `mode` to select the correct segment in the General page and `effective` to set the initial rendering theme.
- **D-12:** `theme_applied(str)` signal continues to emit the effective theme ("dark"/"light") whenever the effective theme changes (system theme polling, explicit mode change). No change to this signal.
- **D-13:** Theme mode selection on the General page routes through `bridge.save_settings` → `apply_draft({"theme": mode})` like all other settings, instead of calling `bridge.apply_theme` directly. The `apply_theme` bridge slot is removed — its logic moves to `_apply_side_effects` for the `theme` key.
- **D-14:** Theme applies immediately when drafted (visual feedback) but is only persisted on commit. On discard, the theme reverts to the previously saved mode. The side-effect in `_apply_side_effects` calls `_apply_theme_mode(applied["theme"])` — already there, just needs to also apply during draft for immediate visual feedback.

### Strict boolean parsing (BRDG-04)
- **D-15:** Add a `_strict_bool(value)` function in `virelo/settings/state.py` that accepts only: `True`, `False`, `"true"`, `"false"`, `1`, `0`. Raises `ValueError` for any other input (including `"yes"`, `"no"`, `"on"`, `"off"`, `None`, non-boolean strings). This prevents `bool("false")` → `True`.
- **D-16:** Replace the `bool` coercer in `SettingsState.KEYS` with `_strict_bool` for all boolean keys (`enable_snap`, `ex_auto_size`, `game_mode_enabled`, `run_at_startup`).
- **D-17:** `_safe_bool` in `persistence.py` remains unchanged — QSettings reads need permissive parsing because the Windows Registry can store values in unpredictable formats. Strict parsing is enforced at the bridge boundary (SettingsState) where frontend values enter Python.

### UI preference persistence (BRDG-06)
- **D-18:** Add three new settings keys to the Python model:
  - `accent` (str): valid values `"slate"`, `"teal"`, `"blue"`, `"rust"`, `"purple"` — default `"slate"`
  - `density` (str): valid values `"compact"`, `"cozy"`, `"comfortable"` — default `"cozy"`
  - `minimize_to_tray` (bool): default `True`
- **D-19:** Add these keys to `SettingsState.KEYS`, `Settings.__init__` (QSettings read), `Settings.save()` (QSettings write), `DEFAULTS` in `config.py`, and `bridgeToState`/`stateToBridge` in `app.jsx`.
- **D-20:** Frontend `setTweaks` for accent and density routes through `app.set(...)` → `bridge.save_settings` → `apply_draft`. On app load, accent/density come from `get_settings` alongside other settings. The `ThemeProvider` tweaks state is initialized from bridge settings, not hardcoded defaults.
- **D-21:** `radius` and `sidebarMode` have no visible UI controls — they remain as frontend constants (`radius: 6`, `sidebarMode: 'full'`) and are NOT added to the Python model.
- **D-22:** `minimize_to_tray` replaces the in-memory `minimize_to_tray_on_exit` flag in MainWindow. On startup, read from persisted settings. The tray menu "Minimize to Tray" toggle updates via draft/commit like other settings.

### Claude's Discretion
- Whether `dirty_changed` emits on every `apply_draft` call or is debounced
- Whether to add accent/density validation as an enum or a simple string check
- Internal refactoring of `_apply_side_effects` to handle the expanded set of side-effect keys
- Whether the tray menu "Run at Startup" checkbox updates optimistically during draft or only on commit
- How to handle theme revert on discard (immediate visual revert vs. fade transition)

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Vision, constraints, out-of-scope features, key decisions
- `.planning/REQUIREMENTS.md` — BRDG-01..06 requirement definitions (Phase 5 scope)
- `CLAUDE.md` — Forbidden changes (no fake controls, no stale naming), known footguns

### Prior phase context
- `.planning/phases/02-security-and-bridge-hardening/02-CONTEXT.md` — Draft state architecture (D-06..D-10), bridge response standardization (D-15..D-17), fake control removal (D-11..D-14)
- `.planning/phases/03-structure-and-quality/03-CONTEXT.md` — Package layout (D-01..D-02), test tiering (D-08..D-10)
- `.planning/phases/04-snap-and-explorer-hardening/04-CONTEXT.md` — HotkeyListener/ShiftSnapRestore separation (D-01..D-03), ExplorerService extraction (D-07..D-10)

### Architecture
- `.planning/codebase/ARCHITECTURE.md` — Current architecture layers and data flow
- `.planning/codebase/CONVENTIONS.md` — Naming, imports, error handling patterns
- `.planning/codebase/STRUCTURE.md` — Current file layout

### Current implementation (files being modified)
- `virelo/bridge/bridge.py` — VireloBridge: signals, slots, `_apply_side_effects` (primary modification target)
- `virelo/settings/state.py` — SettingsState: draft/commit model, KEYS dict, `_strict_bool` addition
- `virelo/settings/persistence.py` — Settings: QSettings read/write, `_safe_bool`/`_safe_int` (add new keys)
- `virelo/app/config.py` — DEFAULTS dict (add `accent`, `density`, `minimize_to_tray`)
- `virelo/app/window.py` — MainWindow: key capture flow, theme handling, startup shortcut, tray menu sync
- `virelo/platform/theme.py` — Theme resolution (unchanged, but referenced by theme coherence work)
- `frontend/src/app.jsx` — VireloApp: `bridgeToState`/`stateToBridge`, dirty state subscription, `set()` handler
- `frontend/src/main.jsx` — Root: initial theme loading, `handleSetTweaks` bridge routing
- `frontend/src/pages.jsx` — GeneralPage: accent/density controls, launch-at-login toggle, theme segmented
- `frontend/src/theme.jsx` — ThemeProvider tweaks, ACCENTS/DENSITIES constants (referenced but not modified)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `SettingsState` draft/commit model — already has `apply_draft()`, `commit_draft()`, `discard_draft()`, `has_draft`. Extend with new keys and `dirty_changed` emission.
- `_apply_side_effects(applied)` in bridge.py — already dispatches on `enable_snap`, `ex_auto_size`, `snap_presses`, `snap_key`, `restore_key`, `theme`. Add `run_at_startup`, `accent`, `density`, `minimize_to_tray`.
- `_safe_bool`/`_safe_int` in persistence.py — defensive coercion for QSettings reads. Pattern to follow for new keys.
- `bridgeToState`/`stateToBridge` in app.jsx — key mapping between Python snake_case and React camelCase. Extend with new keys.
- `create_startup_shortcut()`/`remove_startup_shortcut()` in window.py — standalone functions, ready to call from `_apply_side_effects`.

### Established Patterns
- All bridge slots return `{"ok": true, "data": ...}` / `{"ok": false, "error": "..."}` — new slots and modifications follow this pattern.
- Signal-slot for Python→JS communication — `dirty_changed(bool)` follows `settings_changed(str)` pattern.
- Draft overlay in `get_all()` — persisted values overlaid with draft values, normalized after overlay.
- Side effects fire on commit, not on draft — `_apply_side_effects` called in `commit_draft`, not `apply_draft`.

### Integration Points
- `MainWindow.__init__` constructs Settings → SettingsState → VireloBridge — new keys must be initialized in Settings constructor.
- `settings_changed` signal pushes full settings to frontend after draft/commit/discard — `dirty_changed` emits alongside.
- `main.jsx` `AppWithBridge` initializes tweaks from hardcoded defaults — must initialize from bridge settings instead.
- `GeneralPage` theme/accent/density controls use `setTweaks` — must route through `app.set` for persistence.

</code_context>

<specifics>
## Specific Ideas

- The dirty state signal (D-01) is the foundational change — all other BRDG requirements depend on the frontend trusting Python-driven dirty state rather than local inference.
- Key capture draft integration (D-04..D-07) is the most delicate change: the capture worker runs on a background thread and currently writes directly to settings. The draft path must be thread-safe (SettingsState._draft is accessed from both the main thread and capture result handler, but Qt signal-slot handles the threading).
- Boolean strictness (D-15..D-17) is a small but important correctness fix. The current `bool()` coercer in KEYS silently accepts `"false"` as `True`, which is a real bug that could manifest when the frontend sends string "false" instead of boolean false.
- UI preference persistence (D-18..D-22) closes the loop on BRDG-06. Without this, accent and density reset to defaults every app restart. The user's visual customization is lost.

</specifics>

<deferred>
## Deferred Ideas

None — analysis stayed within phase scope

</deferred>

---

*Phase: 05-bridge-and-settings-correctness*
*Context gathered: 2026-04-24*
