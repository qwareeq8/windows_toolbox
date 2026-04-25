# Phase 5: Bridge and Settings Correctness - Research

**Researched:** 2026-04-24
**Domain:** PySide6/QWebChannel bridge signals, Python settings draft/commit model, React state synchronization
**Confidence:** HIGH

## Summary

Phase 5 corrects the bridge-level contract between the Python backend and React frontend so that every setting flows through a single Python-owned draft/commit model with correct types and coherent signals. The codebase already has the draft/commit model (`SettingsState`), the bridge (`VireloBridge`), and the `_apply_side_effects` dispatch pattern. The work is primarily: (1) adding a `dirty_changed(bool)` signal and wiring the frontend to consume it instead of local `setUnsaved` inference, (2) routing key capture results through `apply_draft` instead of direct writes, (3) routing launch-at-login and theme through `_apply_side_effects` instead of standalone bridge slots, (4) adding `_strict_bool` validation at the bridge boundary, and (5) persisting accent/density/minimize_to_tray through the existing settings model.

The existing patterns are well-established from v1.0 phases: all bridge slots return `{"ok": true, "data": ...}` / `{"ok": false, "error": "..."}`, signals push from Python to JS, draft accumulates until commit, side effects fire on commit. The changes in this phase are extensions of these patterns, not new architectural patterns. The primary risk areas are: (a) the key capture callback `_on_capture_key` runs via Qt signal-slot from a background thread, so the draft write it needs to make is automatically marshalled to the main thread -- but the implementer must not add any threading beyond the existing pattern; (b) theme must apply visually during draft for immediate feedback (D-14), which is the one exception to the "side effects on commit only" rule; (c) the `set()` function in app.jsx sends the FULL state on every change via `stateToBridge(next)`, so after adding accent/density/minimize_to_tray to the mapping, every setting change will include these keys in the draft payload.

**Primary recommendation:** Implement in logical layers -- first the `dirty_changed` signal foundation (all other features depend on it), then the strict boolean fix (small scope, validates the draft path correctness), then key capture via draft (most delicate thread-safety concern), then launch-at-login and theme coherence (both reroute existing features through draft), then UI preference persistence (extends the model with new keys).

<user_constraints>

## User Constraints (from CONTEXT.md)

### Locked Decisions

- **D-01:** Add a `dirty_changed(bool)` signal on `VireloBridge`. Emit `True` when `apply_draft()` creates or extends a draft, `False` when `commit_draft()` or `discard_draft()` clears it. Also emit `False` after `reset_to_defaults()`.
- **D-02:** Frontend replaces its local `unsaved` React state with a subscription to `dirty_changed`. The `setUnsaved(true)` call in `app.jsx:set()` and the `setUnsaved(false)` calls in `handleSave`/`handleDiscard`/`handleReset` are removed -- the signal drives the footer dirty indicator.
- **D-03:** The `settings_changed` signal continues to push the full settings dict (including draft overlay). `dirty_changed` is a separate signal because dirty state is orthogonal to settings values.
- **D-04:** `_on_capture_key` in `window.py` routes the captured key through `self._settings_state.apply_draft({"snap_key": key_str})` (or `"restore_key"`) instead of writing directly to `self.settings.snap_key`. This triggers `dirty_changed(True)`.
- **D-05:** The captured key does NOT immediately update the `HotkeyListener` binding. The old binding stays active until `commit_draft`. Hotkey listener updates happen in `_apply_side_effects` on commit.
- **D-06:** On discard, the key reverts to the previously saved value. `discard_draft()` clears `_draft`, and `settings_changed` pushes persisted values.
- **D-07:** The `on_key_captured` slot in `MainWindow` that directly sets `self.settings.snap_key` and emits `snap_status` is removed. All capture result handling goes through `_on_capture_key` which now uses the draft path.
- **D-08:** The `toggle_run_at_startup` bridge slot is removed. Launch at login changes flow through the normal `save_settings` -> `apply_draft` path. The General page's "Launch at login" toggle calls `app.set({ launchLogin: v })`.
- **D-09:** `_apply_side_effects` gains a `run_at_startup` handler: when `run_at_startup` is in the committed dict, call `create_startup_shortcut()` or `remove_startup_shortcut()` accordingly. Wrap in try/except and report errors via `snap_status` signal.
- **D-10:** The tray menu "Run at Startup" action syncs with the persisted setting on startup and after each `commit_draft`. It does NOT directly toggle the shortcut -- it goes through the bridge.
- **D-11:** `get_theme_mode` slot returns JSON `{"mode": "system", "effective": "dark"}` -- both the user's chosen mode and the resolved effective theme.
- **D-12:** `theme_applied(str)` signal continues to emit the effective theme ("dark"/"light") whenever the effective theme changes.
- **D-13:** Theme mode selection routes through `bridge.save_settings` -> `apply_draft({"theme": mode})` instead of calling `bridge.apply_theme` directly. The `apply_theme` bridge slot is removed.
- **D-14:** Theme applies immediately when drafted (visual feedback) but is only persisted on commit. On discard, the theme reverts.
- **D-15:** Add a `_strict_bool(value)` function in `virelo/settings/state.py` that accepts only: `True`, `False`, `"true"`, `"false"`, `1`, `0`. Raises `ValueError` for any other input.
- **D-16:** Replace the `bool` coercer in `SettingsState.KEYS` with `_strict_bool` for all boolean keys.
- **D-17:** `_safe_bool` in `persistence.py` remains unchanged -- QSettings reads need permissive parsing.
- **D-18:** Add three new settings keys: `accent` (str), `density` (str), `minimize_to_tray` (bool).
- **D-19:** Add these keys to `SettingsState.KEYS`, `Settings.__init__`, `Settings.save()`, `DEFAULTS`, and `bridgeToState`/`stateToBridge`.
- **D-20:** Frontend `setTweaks` for accent and density routes through `app.set(...)` -> `bridge.save_settings` -> `apply_draft`. On app load, accent/density come from `get_settings`.
- **D-21:** `radius` and `sidebarMode` remain as frontend constants, NOT added to Python model.
- **D-22:** `minimize_to_tray` replaces the in-memory `minimize_to_tray_on_exit` flag in MainWindow.

### Claude's Discretion

- Whether `dirty_changed` emits on every `apply_draft` call or is debounced
- Whether to add accent/density validation as an enum or a simple string check
- Internal refactoring of `_apply_side_effects` to handle the expanded set of side-effect keys
- Whether the tray menu "Run at Startup" checkbox updates optimistically during draft or only on commit
- How to handle theme revert on discard (immediate visual revert vs. fade transition)

### Deferred Ideas (OUT OF SCOPE)

None -- analysis stayed within phase scope

</user_constraints>

<phase_requirements>

## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| BRDG-01 | User sees accurate dirty/clean state in the footer driven by Python dirty_changed signal, not local React inference | D-01/D-02/D-03: `dirty_changed(bool)` signal on VireloBridge, frontend subscribes instead of local `setUnsaved`. Existing `Signal(bool)` pattern verified in PySide6 docs. |
| BRDG-02 | User can toggle Launch at login and have the startup shortcut created or removed on Save, with error reporting on failure | D-08/D-09/D-10: Remove `toggle_run_at_startup` slot, route through draft/commit, add `run_at_startup` handler in `_apply_side_effects`. Existing `create_startup_shortcut()`/`remove_startup_shortcut()` functions ready to use. |
| BRDG-03 | User can capture a new key binding and see it reflected as a dirty draft change before saving | D-04/D-05/D-06/D-07: Route `_on_capture_key` through `apply_draft` instead of direct write. Remove `on_key_captured` slot. Thread safety handled by Qt signal-slot marshalling. |
| BRDG-04 | User cannot cause silent boolean coercion bugs through bridge settings (strict parsing of true/false/1/0) | D-15/D-16/D-17: Add `_strict_bool()` function, replace `bool` coercer in `KEYS` dict for boolean keys. `_safe_bool` in persistence.py stays permissive. |
| BRDG-05 | User can select System, Light, or Dark theme with Python sending both theme_mode and effective_theme to the frontend | D-11/D-12/D-13/D-14: Remove `apply_theme` slot, update `get_theme_mode` to return `{mode, effective}`, route theme through draft/commit with immediate visual application. Add "System" to frontend theme selector. |
| BRDG-06 | User's UI preferences (accent, density, minimize-to-tray) are persisted through the Python draft model | D-18/D-19/D-20/D-21/D-22: Add `accent`, `density`, `minimize_to_tray` to DEFAULTS, Settings, SettingsState, and bridge key mappings. Frontend loads tweaks from bridge on startup instead of hardcoding. |

</phase_requirements>

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Dirty state tracking | Python (VireloBridge) | Frontend (subscription) | Python owns the single source of truth for draft state; frontend only reflects it via signal |
| Settings persistence | Python (Settings/QSettings) | -- | QSettings writes to Windows Registry; no frontend involvement in persistence |
| Draft/commit model | Python (SettingsState) | -- | All validation, coercion, and draft accumulation is Python-side |
| Key capture | Python (KeyCaptureWorker on QThread) | Frontend (status display) | Background thread for keyboard hook; results flow through Python bridge |
| Theme resolution | Python (theme.py) | Frontend (visual rendering) | Python resolves "system" to effective theme via registry read; frontend applies CSS tokens |
| Startup shortcut | Python (win32com/WScript.Shell) | -- | OS-level shortcut manipulation is entirely Python-side |
| UI preference display | Frontend (ThemeProvider) | Python (persistence) | Frontend renders accent/density; Python stores values |
| Tray menu sync | Python (MainWindow) | -- | Tray menu is Qt-native; updated from Python after commit |

## Standard Stack

### Core (already established -- no new dependencies)

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| PySide6 | (project-installed) | QWebChannel bridge, signals, QSettings | Established stack, not changing |
| React | 19 | Frontend UI rendering | Established stack, not changing |
| Vite | (project-installed) | Frontend build tool | Established stack, not changing |

No new libraries are needed for this phase. All work extends existing patterns within the established PySide6 + React + QWebChannel stack. [VERIFIED: codebase inspection]

## Architecture Patterns

### System Architecture Diagram

```
Frontend (React)                          Python Backend (PySide6)
=================                         =======================

User interacts with                       VireloBridge (QObject)
settings controls                         +-- dirty_changed(bool) signal
       |                                  +-- settings_changed(str) signal
       v                                  +-- theme_applied(str) signal
app.set({key: val})                        +-- capture_status(str) signal
       |                                  +-- snap_status(str, int) signal
       v                                        |
stateToBridge(next) ----[QWebChannel]----> save_settings(json_str)
       |                                        |
       |                                        v
       |                                  SettingsState.apply_draft(data)
       |                                        |
       |                                        +-- validates & coerces
       |                                        +-- stores in _draft
       |                                        +-- Bridge emits dirty_changed(True)
       |                                        +-- Bridge emits settings_changed(json)
       |                                        |
       |                                  [for theme only: immediate visual apply]
       |                                        |
       v                                        v
"Save" button --------[QWebChannel]----> commit_draft()
       |                                        |
       |                                        v
       |                                  SettingsState.commit_draft()
       |                                        +-- writes to Settings attrs
       |                                        +-- Settings.save() -> QSettings -> Registry
       |                                        +-- Bridge emits dirty_changed(False)
       |                                        +-- Bridge emits settings_changed(json)
       |                                        +-- _apply_side_effects(applied)
       |                                             +-- snap_key -> HotkeyListener
       |                                             +-- theme -> _apply_theme_mode
       |                                             +-- run_at_startup -> shortcut
       |                                             +-- minimize_to_tray -> flag
       |                                             +-- sync tray menu checkboxes
       |                                        
       v                                        
"Discard" button -----[QWebChannel]----> discard_draft()
                                                +-- clears _draft
                                                +-- Bridge emits dirty_changed(False)
                                                +-- Bridge emits settings_changed(json)
                                                +-- [theme revert: _apply_theme_mode(persisted)]
```

### Key Capture Data Flow (revised per D-04..D-07)

```
Frontend                      MainWindow                    KeyCaptureWorker (QThread)
--------                      ----------                    -------------------------
capture_key("snap") ------>  _begin_key_capture()
                                   |
                                   +-- CaptureGuard.try_start()
                                   +-- create QThread + worker
                                   +-- worker.captured.connect(_on_capture_key)
                                   |
                                   |                        keyboard.hook() -> polls
                                   |                                |
                                   |                        user presses key
                                   |                                |
                                   |         <-- captured signal -- worker emits captured(key)
                                   |
                              _on_capture_key(key)
                                   |
                                   +-- apply_draft({"snap_key": key})  [NOT direct write]
                                   +-- dirty_changed(True) emitted
                                   +-- settings_changed emitted (shows new key in UI)
                                   +-- capture_status("done") emitted
                                   |
                                   +-- Old hotkey binding STAYS active until commit
```

### Pattern 1: dirty_changed Signal

**What:** A `Signal(bool)` on VireloBridge that the frontend subscribes to for dirty state tracking.
**When to use:** Emitted by the bridge methods that modify draft state: `save_settings` (after `apply_draft`), `commit_draft`, `discard_draft`, `reset_defaults`.

```python
# Source: PySide6 Signal docs (https://doc.qt.io/qtforpython-6/tutorials/basictutorial/signals_and_slots.html)
# and existing settings_changed pattern in bridge.py

class VireloBridge(QObject):
    dirty_changed = Signal(bool)  # True = has unsaved draft, False = clean
    
    # In save_settings, after apply_draft succeeds:
    def save_settings(self, json_str):
        # ... existing validation ...
        result = self._state.apply_draft(data)
        if result.get("ok"):
            self.settings_changed.emit(self._state.get_json())
            self.dirty_changed.emit(self._state.has_draft)
        return json.dumps(result)
    
    # In commit_draft, after commit succeeds:
    def commit_draft(self):
        result = self._state.commit_draft()
        if result.get("ok"):
            self.settings_changed.emit(self._state.get_json())
            self.dirty_changed.emit(False)  # draft cleared
            self._apply_side_effects(result.get("applied", {}))
        return json.dumps(result)
```

[VERIFIED: PySide6 Signal(bool) is standard pattern -- confirmed via Context7 docs at doc.qt.io/qtforpython-6/tutorials/basictutorial/signals_and_slots.html]

### Pattern 2: _strict_bool Validation

**What:** A strict boolean coercer that only accepts true boolean-like values, rejecting ambiguous inputs.
**When to use:** Replaces `bool` in `SettingsState.KEYS` for boolean setting keys.

```python
# In virelo/settings/state.py

def _strict_bool(value):
    """Parse strict boolean values. Raises ValueError for ambiguous input.
    
    Accepts: True, False, "true", "false", 1, 0
    Rejects: "yes", "no", "on", "off", None, non-boolean strings
    
    This prevents Python's bool("false") == True pitfall at the bridge boundary.
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        if value in (0, 1):
            return bool(value)
        raise ValueError(f"Expected 0 or 1, got {value}")
    if isinstance(value, str):
        lower = value.strip().lower()
        if lower == "true":
            return True
        if lower == "false":
            return False
        raise ValueError(f"Expected 'true' or 'false', got '{value}'")
    raise ValueError(f"Cannot convert {type(value).__name__} to bool")
```

[VERIFIED: codebase shows `bool` currently used as coercer in KEYS dict -- `bool("false")` returns `True` in Python, confirming the bug]

### Pattern 3: Theme Immediate Apply on Draft (exception to side-effects-on-commit)

**What:** Theme mode applies visually immediately when drafted, unlike other settings which only take effect on commit.
**When to use:** When user selects System/Light/Dark in the theme picker.

```python
# In VireloBridge.save_settings, after apply_draft:
if result.get("ok"):
    self.settings_changed.emit(self._state.get_json())
    self.dirty_changed.emit(self._state.has_draft)
    # Theme is special: apply immediately for visual feedback (D-14)
    if "theme" in data:
        theme_mode = result["applied"].get("theme")
        if theme_mode and self._main_window:
            self._main_window._apply_theme_mode(theme_mode)
```

### Pattern 4: Side Effect Dispatch Expansion

**What:** Extending `_apply_side_effects` with new handlers for `run_at_startup`, `minimize_to_tray`, and tray menu sync.
**When to use:** Called after `commit_draft` returns successfully.

```python
# Extend _apply_side_effects in bridge.py

def _apply_side_effects(self, applied: dict):
    if not self._main_window:
        return
    mw = self._main_window
    
    # ... existing handlers for enable_snap, ex_auto_size, snap_presses, snap_key, restore_key, theme ...
    
    if "run_at_startup" in applied:
        try:
            if applied["run_at_startup"]:
                create_startup_shortcut()
            else:
                remove_startup_shortcut()
        except Exception as e:
            LOG.exception("Startup shortcut error")
            self.snap_status.emit(f"Failed to update startup shortcut: {e}", 5000)
        # Sync tray menu checkbox
        mw.action_run_at_startup.setChecked(bool(applied["run_at_startup"]))
    
    if "minimize_to_tray" in applied:
        mw.minimize_to_tray_on_exit = bool(applied["minimize_to_tray"])
        mw.action_minimize_on_exit.setChecked(bool(applied["minimize_to_tray"]))
```

### Pattern 5: Frontend bridgeToState/stateToBridge Extension

**What:** Adding accent, density, minimize_to_tray to the key mapping functions.
**When to use:** When the frontend loads settings or sends updates to the bridge.

```javascript
// In app.jsx

export function bridgeToState(settings) {
  return {
    // ... existing keys ...
    accent: settings.accent || 'slate',
    density: settings.density || 'cozy',
    minimizeToTray: settings.minimize_to_tray ?? true,
  };
}

export function stateToBridge(state) {
  return JSON.stringify({
    // ... existing keys ...
    accent: state.accent,
    density: state.density,
    minimize_to_tray: state.minimizeToTray,
  });
}
```

### Anti-Patterns to Avoid

- **Direct settings write from key capture:** The current `_on_capture_key` writes `self.settings.snap_key = key_str` directly. This bypasses the draft model and makes the key take effect before the user saves. Must route through `apply_draft`.
- **Standalone bridge slots for individual settings:** The current `toggle_run_at_startup` and `apply_theme` slots bypass the draft/commit flow. Must be removed; all settings go through `save_settings` -> `apply_draft`.
- **Local unsaved state inference:** The frontend currently infers dirty state (`setUnsaved(true)` in `set()`, `setUnsaved(false)` in callbacks). Must be replaced with signal subscription.
- **Python `bool()` as type coercer:** `bool("false")` returns `True` in Python. Must use `_strict_bool` for bridge boundary validation.
- **Hardcoded frontend defaults for persisted settings:** `AppWithBridge` currently hardcodes `accent: 'slate'`, `density: 'cozy'`. Must load from bridge `get_settings`.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Boolean parsing | Custom truthy/falsy logic | `_strict_bool()` as defined in D-15 | Python's `bool()` has the well-known "false" -> True trap; a strict parser eliminates ambiguity |
| Thread-safe signal delivery | Manual locking for capture results | Qt signal-slot automatic thread marshalling | Qt automatically queues signals across threads when connected between QObjects on different threads |
| Theme resolution | Manual OS theme detection | Existing `resolve_theme()` + `get_windows_theme()` in platform/theme.py | Already handles registry reads and normalization |
| Startup shortcut | Manual .lnk creation | Existing `create_startup_shortcut()` / `remove_startup_shortcut()` in window.py | Already handles frozen vs. dev mode, icon paths, gen_py cache corruption |

**Key insight:** This phase extends existing infrastructure, not building new frameworks. Every feature uses established patterns (draft/commit, signal-slot, side-effect dispatch). The work is wiring and correctness, not architecture.

## Common Pitfalls

### Pitfall 1: settings_changed Resets Unsaved State

**What goes wrong:** The current `settings_changed.connect` handler in app.jsx calls `setUnsaved(false)`. After D-02 removes `setUnsaved`, if the frontend still reacts to `settings_changed` by clearing dirty state, it will conflict with the `dirty_changed` signal.
**Why it happens:** `settings_changed` fires after both `apply_draft` (where dirty should become true) AND `commit_draft`/`discard_draft` (where dirty should become false). Using it for dirty tracking is inherently wrong.
**How to avoid:** Remove the `setUnsaved(false)` from the `settings_changed` handler. Use `dirty_changed` exclusively for dirty state.
**Warning signs:** Dirty indicator flickers or fails to appear after a setting change.

### Pitfall 2: bool("false") == True

**What goes wrong:** Frontend sends `"false"` as a string (possible with JSON serialization edge cases). Python `bool("false")` returns `True` because non-empty strings are truthy.
**Why it happens:** QWebChannel serializes JavaScript values. While `JSON.stringify` should preserve booleans, the coercion chain in `SettingsState.KEYS` uses bare `bool` which accepts any truthy value.
**How to avoid:** Replace `bool` with `_strict_bool` in KEYS for all boolean keys (D-15/D-16).
**Warning signs:** A setting that was toggled off appears on after save/reload.

### Pitfall 3: Key Capture Thread Safety

**What goes wrong:** `_on_capture_key` is connected to `KeyCaptureWorker.captured` signal. The worker runs on a QThread. If the handler accesses `_draft` directly or calls non-thread-safe methods, corruption could occur.
**Why it happens:** Qt signal-slot with queued connection (default for cross-thread) ensures the slot runs on the receiver's thread (main thread). But if someone adds direct method calls to the worker or bypasses signal-slot, thread safety breaks.
**How to avoid:** Keep all draft modification in the main thread via signal-slot. The existing pattern already handles this: `self._capture_worker.captured.connect(self._on_capture_key)` uses a queued connection because the worker is on a different thread.
**Warning signs:** Intermittent crashes or corrupted draft state after key capture.

### Pitfall 4: Theme Revert on Discard

**What goes wrong:** User changes theme from Dark to Light (visual preview applied), changes another setting, then discards. The theme must revert to Dark, but if `discard_draft` only clears `_draft` without re-applying the persisted theme, the UI stays Light.
**Why it happens:** `discard_draft` in SettingsState just sets `_draft = None`. It doesn't trigger side effects. The bridge's `discard_draft` slot needs to additionally revert the theme.
**How to avoid:** In the bridge's `discard_draft` method, after clearing the draft, check if the persisted theme differs from the current visual theme and call `_apply_theme_mode` with the persisted value.
**Warning signs:** Theme stays in the draft-preview state after discard.

### Pitfall 5: Full State Sends on Every Change

**What goes wrong:** The frontend `set()` function calls `stateToBridge(next)` which serializes ALL settings, not just the changed one. After adding accent/density/minimize_to_tray to the mapping, every single slider drag or toggle sends all 14+ keys to `apply_draft`.
**Why it happens:** `stateToBridge` builds the complete settings object and `save_settings` applies it all as a draft.
**How to avoid:** This is a known characteristic, not a bug. `apply_draft` is idempotent for unchanged values. Performance impact is negligible for the small number of keys. Do not attempt to add partial-update optimization in this phase.
**Warning signs:** None -- this works fine but the planner should be aware.

### Pitfall 6: Tray Menu Desync

**What goes wrong:** The tray menu "Run at Startup" checkbox shows a different state than the settings UI.
**Why it happens:** Currently the tray menu toggles the shortcut directly via `_toggle_run_at_startup`. After this phase, the tray menu must route through the bridge like the frontend. But the tray menu action is Qt-native, not React -- it can't call `bridge.save_settings`.
**How to avoid:** The tray menu checkbox should be updated in `_apply_side_effects` after `commit_draft`, reading the committed value. The tray action handler should either: (a) call bridge methods programmatically, or (b) sync the checkbox to match persisted state on startup and after each commit. Per D-10, it should go through the bridge like the frontend control.
**Warning signs:** Tray says "Run at Startup" is checked but the setting is not persisted, or vice versa.

## Code Examples

### Example 1: dirty_changed Frontend Subscription

```javascript
// In app.jsx, inside the useEffect that sets up bridge subscriptions:

// Subscribe to Python-driven dirty state (replaces local setUnsaved)
bridge.dirty_changed.connect((isDirty) => {
  setUnsaved(isDirty);
});

// Remove setUnsaved(true) from set() function
// Remove setUnsaved(false) from handleSave, handleDiscard, handleReset callbacks
// Remove setUnsaved(false) from settings_changed handler
```

[VERIFIED: QWebChannel signal subscription pattern confirmed in PySide6 docs -- `channel.objects.bridge.mySignal.connect(function(args) { ... })` is the standard approach]

### Example 2: get_theme_mode Returning Both Mode and Effective

```python
@Slot(result=str)
def get_theme_mode(self) -> str:
    """Return current theme mode and effective theme."""
    mode = "system"
    effective = "dark"
    if self._main_window:
        mode = getattr(self._main_window, "_theme_mode", "system")
        effective = getattr(self._main_window, "_theme_state", "dark")
    return json.dumps({"ok": True, "data": {"mode": mode, "effective": effective}})
```

### Example 3: Frontend Theme Initialization with Mode + Effective

```javascript
// In main.jsx Root component:
b.get_theme_mode((result) => {
  try {
    const r = JSON.parse(result);
    if (r.ok && r.data) {
      const { mode, effective } = r.data;
      setBridgeState({
        bridge: b,
        initialTheme: effective || 'dark',
        initialThemeMode: mode || 'system',
      });
    }
  } catch (e) {
    setBridgeState({ bridge: b, initialTheme: 'dark', initialThemeMode: 'system' });
  }
});
```

### Example 4: Mock Bridge Update for Dev Mode

```javascript
// In bridge.js, update MOCK_BRIDGE and MOCK_SETTINGS:

const MOCK_SETTINGS = {
  // ... existing keys ...
  accent: 'slate',
  density: 'cozy',
  minimize_to_tray: true,
};

const MOCK_BRIDGE = {
  // ... existing methods ...
  // Remove: apply_theme, toggle_run_at_startup
  // Update: get_theme_mode returns {mode, effective}
  get_theme_mode: (cb) => cb(JSON.stringify({ ok: true, data: { mode: 'dark', effective: 'dark' } })),
  // Add: dirty_changed signal
  dirty_changed: { connect: () => {} },
};
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Standalone bridge slots per setting (`toggle_run_at_startup`, `apply_theme`) | All settings through unified draft/commit via `save_settings` | This phase (D-08, D-13) | Eliminates bypass paths; single flow for all settings |
| Frontend-inferred dirty state (`setUnsaved` local React state) | Python-driven dirty state via `dirty_changed(bool)` signal | This phase (D-01, D-02) | Single source of truth; no UI/backend disagreement |
| Direct key write from capture (`self.settings.snap_key = key`) | Capture routes through `apply_draft` | This phase (D-04, D-07) | Captured key shown as pending; old binding stays until save |
| Python `bool()` coercion for bridge input | `_strict_bool()` strict parsing | This phase (D-15, D-16) | Prevents `bool("false") == True` bug |
| Hardcoded frontend tweaks (accent, density) | Persisted through Python settings model | This phase (D-18, D-20) | UI preferences survive app restart |

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | Qt queued connection is automatic when signal and slot are on different threads (worker thread -> main thread) | Common Pitfalls: Key Capture Thread Safety | If not automatic, `_on_capture_key` would need explicit `Qt.QueuedConnection` in `.connect()` call -- low risk since this is standard Qt behavior |

**Note:** A1 is standard Qt behavior documented extensively. The existing codebase already relies on this for `captured.connect(self._on_capture_key)` where the worker is on a QThread. Risk of this assumption being wrong is negligible.

## Open Questions

1. **Tray menu routing through bridge**
   - What we know: D-10 says tray menu "goes through the bridge like the frontend control." But the tray menu is a Qt QAction, not a web control. It cannot call `bridge.save_settings()` with a JSON string via QWebChannel.
   - What's unclear: How exactly should the tray menu "Run at Startup" toggle interact with the draft/commit model? Should it call `self._settings_state.apply_draft({"run_at_startup": checked})` directly, then `self._settings_state.commit_draft()` and `self._bridge._apply_side_effects(applied)` in sequence? Or should it be a one-step operation that drafts and commits immediately?
   - Recommendation: The tray menu should call the same Python methods that the bridge slots wrap: `apply_draft` then `commit_draft`. This avoids a separate code path while keeping the tray menu working without QWebChannel. The tray handler calls `self._settings_state.apply_draft(...)`, then `self._settings_state.commit_draft()`, then `self._bridge._apply_side_effects(applied)`, and emits both signals. This is Claude's discretion per CONTEXT.md.

2. **Theme selector "System" option in frontend**
   - What we know: The current theme Segmented in GeneralPage only has Light/Dark options. D-11 and D-13 require System/Light/Dark support.
   - What's unclear: The theme selector currently lives in GeneralPage and calls `setTweaks({ theme: v })`. After D-13, it needs to call `app.set({ theme: v })` instead. But `tweaks.theme` is the *effective* theme ("dark"/"light"), not the mode ("system"/"dark"/"light"). The Segmented value needs to track the mode, while `ThemeProvider` uses the effective theme.
   - Recommendation: Add a `themeMode` field to the app state (via `bridgeToState`) that tracks the user-selected mode. The Segmented uses `themeMode` as its value. The `ThemeProvider` continues to receive the effective theme. When the user selects "System", `app.set({ themeMode: 'system' })` sends `theme: 'system'` to Python. Python resolves and emits `theme_applied("dark")` or `theme_applied("light")` for visual rendering.

## Discretion Recommendations

Based on the Claude's Discretion items in CONTEXT.md:

1. **dirty_changed debouncing:** Emit on every `apply_draft` call without debounce. The signal is lightweight (single bool) and the frontend React re-render of the dirty indicator is trivial. Debouncing would add complexity with no measurable benefit. [ASSUMED: based on React rendering cost being negligible for a single boolean state update]

2. **accent/density validation:** Use a simple `in` check against a tuple of valid values, not an enum class. This matches the project's existing validation style (e.g., `normalize_theme_mode` uses string comparison). Add a helper function:
   ```python
   def _validate_enum(value, allowed, default):
       v = str(value).strip().lower()
       return v if v in allowed else default
   ```

3. **_apply_side_effects refactoring:** Keep the existing if-chain pattern. The method is straightforward and readable. Adding a dispatch table or strategy pattern would be over-engineering for 8-10 simple handlers.

4. **Tray menu "Run at Startup" update timing:** Update on commit only, not optimistically during draft. The tray menu reflects persisted state. If the user drafts a change but discards, the tray menu should never have flickered.

5. **Theme revert on discard:** Immediate visual revert, no fade transition. The app does not use CSS transitions for theme changes currently, and adding a fade would be scope creep. Just call `_apply_theme_mode(persisted_theme)` in the bridge's `discard_draft` handler.

## Project Constraints (from CLAUDE.md)

- **Never add fake or placeholder UI controls that do not connect to real backend logic.** All controls added in this phase (System theme option, accent/density persistence) must be fully wired to Python bridge.
- **Never hardcode version strings.** Use `APP_VERSION` from `virelo/app/config.py`.
- **Naming conventions:** Python `snake_case`, JS `camelCase`, React `PascalCase`.
- **Bridge response format:** All slots return `{"ok": true, "data": ...}` / `{"ok": false, "error": "..."}`.
- **Forbidden:** "Windows Toolbox" or "Toolbox" naming in any file.

## Sources

### Primary (HIGH confidence)
- Codebase inspection: `virelo/bridge/bridge.py`, `virelo/settings/state.py`, `virelo/settings/persistence.py`, `virelo/app/config.py`, `virelo/app/window.py`, `frontend/src/app.jsx`, `frontend/src/main.jsx`, `frontend/src/pages.jsx`, `frontend/src/theme.jsx`, `frontend/src/bridge.js`
- Context7 `/websites/doc_qt_io_qtforpython-6` - PySide6 Signal declaration, QWebChannel JavaScript usage, signal-slot patterns
- CONTEXT.md decisions D-01 through D-22

### Secondary (MEDIUM confidence)
- PySide6 Qt documentation on signal-slot thread marshalling behavior

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - No new dependencies; all work extends existing patterns verified in codebase
- Architecture: HIGH - All patterns are extensions of existing draft/commit and signal-slot infrastructure; no new architectural decisions needed
- Pitfalls: HIGH - All pitfalls identified from direct codebase inspection of current code that will be modified

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (30 days - stable domain, no external dependency changes)
