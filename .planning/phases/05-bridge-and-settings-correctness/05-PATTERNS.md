# Phase 5: Bridge and Settings Correctness - Pattern Map

**Mapped:** 2026-04-24
**Files analyzed:** 10 (8 modified, 2 reference-only)
**Analogs found:** 8 / 8

## File Classification

| New/Modified File | Role | Data Flow | Closest Analog | Match Quality |
|-------------------|------|-----------|----------------|---------------|
| `virelo/bridge/bridge.py` | bridge | request-response | self (extending existing) | exact |
| `virelo/settings/state.py` | model | transform | self (extending existing) | exact |
| `virelo/settings/persistence.py` | model | CRUD | self (extending existing) | exact |
| `virelo/app/config.py` | config | static | self (extending existing) | exact |
| `virelo/app/window.py` | controller | event-driven | self (refactoring existing) | exact |
| `frontend/src/app.jsx` | component | request-response | self (extending existing) | exact |
| `frontend/src/main.jsx` | component | event-driven | self (refactoring existing) | exact |
| `frontend/src/pages.jsx` | component | event-driven | `frontend/src/app.jsx` | role-match |

## Pattern Assignments

### `virelo/bridge/bridge.py` (bridge, request-response)

**Analog:** Self -- all modifications extend existing patterns in this file.

**Signal declaration pattern** (lines 36-39):
```python
# --- Signals (Python -> JS) ---
settings_changed = Signal(str)  # JSON string of full settings dict
theme_applied = Signal(str)  # "dark" or "light" (effective theme)
snap_status = Signal(str, int)  # (message, timeout_ms)
capture_status = Signal(str)  # "capturing", "done", "cancelled", "timeout"
```
New `dirty_changed = Signal(bool)` follows this exact pattern: class-level Signal declaration with type annotation comment.

**Slot return format** (lines 66-73):
```python
@Slot(result=str)
def get_settings(self) -> str:
    """Return all settings as a structured JSON payload."""
    try:
        settings = self._state.get_all()
        return json.dumps({"ok": True, "data": settings})
    except Exception as e:
        LOG.exception("get_settings failed")
        return json.dumps({"ok": False, "error": str(e)})
```
All slots return `{"ok": true, "data": ...}` or `{"ok": false, "error": "..."}`. Every new or modified slot must follow this.

**save_settings slot -- draft + signal emission** (lines 75-95):
```python
@Slot(str, result=str)
def save_settings(self, json_str: str) -> str:
    try:
        data = json.loads(json_str)
        if not isinstance(data, dict):
            return json.dumps({"ok": False, "error": "Expected JSON object"})
        result = self._state.apply_draft(data)
        if result.get("ok"):
            # Push updated settings (with draft overlay) to frontend
            self.settings_changed.emit(self._state.get_json())
        return json.dumps(result)
    except json.JSONDecodeError as e:
        return json.dumps({"ok": False, "error": f"Invalid JSON: {e}"})
    except Exception as e:
        LOG.exception("save_settings failed")
        return json.dumps({"ok": False, "error": str(e)})
```
After `apply_draft` succeeds, add `self.dirty_changed.emit(self._state.has_draft)` alongside the `settings_changed` emit. Same pattern applies to `commit_draft` (emit `False`) and `discard_draft` (emit `False`).

**commit_draft slot -- side effects dispatch** (lines 97-108):
```python
@Slot(result=str)
def commit_draft(self) -> str:
    try:
        result = self._state.commit_draft()
        if result.get("ok"):
            self.settings_changed.emit(self._state.get_json())
            self._apply_side_effects(result.get("applied", {}))
        return json.dumps(result)
    except Exception as e:
        LOG.exception("commit_draft failed")
        return json.dumps({"ok": False, "error": str(e)})
```
Add `self.dirty_changed.emit(False)` after `settings_changed.emit`.

**_apply_side_effects -- if-chain dispatch** (lines 260-283):
```python
def _apply_side_effects(self, applied: dict):
    """Apply business logic side effects after settings are committed."""
    if not self._main_window:
        return
    mw = self._main_window

    if "enable_snap" in applied:
        mw.snap_enabled = bool(applied["enable_snap"])
        mw._update_snap_enabled_state()

    if "ex_auto_size" in applied:
        mw._update_explorer_autosize_thread()

    if "snap_presses" in applied and hasattr(mw, "_hotkey_listener"):
        mw._hotkey_listener.update_press_limit(applied["snap_presses"])

    if "snap_key" in applied and hasattr(mw, "_hotkey_listener"):
        mw._hotkey_listener.update_binding(applied["snap_key"])

    if "restore_key" in applied and hasattr(mw, "_hotkey_listener"):
        mw._hotkey_listener.update_restore_key(applied["restore_key"])

    if "theme" in applied:
        mw._apply_theme_mode(applied["theme"])
```
New handlers for `run_at_startup`, `minimize_to_tray` follow this same if-chain pattern. Each checks `"key" in applied`, then acts on the value. Error-prone operations (startup shortcut) wrap in try/except.

**Slots to remove** (lines 183-221):
- `apply_theme` (lines 183-197): Logic moves to `_apply_side_effects` for theme key and immediate-apply in `save_settings` for draft preview.
- `toggle_run_at_startup` (lines 209-221): Logic moves to `_apply_side_effects` for `run_at_startup` key.

**get_theme_mode -- current pattern** (lines 199-205):
```python
@Slot(result=str)
def get_theme_mode(self) -> str:
    """Return current theme mode as structured JSON."""
    mode = "dark"
    if self._main_window:
        mode = getattr(self._main_window, "_theme_mode", "dark")
    return json.dumps({"ok": True, "data": mode})
```
Must be updated to return `{"ok": True, "data": {"mode": "system", "effective": "dark"}}` -- both user mode and resolved effective theme. Uses `_theme_mode` for mode and `_theme_state` for effective.

---

### `virelo/settings/state.py` (model, transform)

**Analog:** Self -- extending the KEYS dict and adding `_strict_bool`.

**KEYS dict pattern** (lines 29-41):
```python
KEYS = {
    "snap_key": (str, None),
    "restore_key": (str, None),
    "enable_snap": (bool, None),
    "snap_presses": (int, (1, 10)),
    "snap_interval": (int, (100, 5000)),
    "width_pct": (int, (10, 100)),
    "height_pct": (int, (10, 100)),
    "ex_auto_size": (bool, None),
    "game_mode_enabled": (bool, None),
    "run_at_startup": (bool, None),
    "theme": (str, None),
}
```
Each entry is `"key": (type_coercer, validator_range_or_None)`. New keys follow this exact format:
- `"accent": (str, None)` -- with custom validation in `apply_draft`
- `"density": (str, None)` -- with custom validation in `apply_draft`
- `"minimize_to_tray": (_strict_bool, None)` -- boolean with strict coercion
- Replace all `bool` coercers with `_strict_bool`: `enable_snap`, `ex_auto_size`, `game_mode_enabled`, `run_at_startup`

**apply_draft -- coercion + validation** (lines 76-112):
```python
def apply_draft(self, data: dict) -> dict:
    unknown = [k for k in data if k not in self.KEYS]
    if unknown:
        return {"ok": False, "error": f"Unknown keys: {unknown}"}

    validated = {}
    for key, value in data.items():
        coercer, bounds = self.KEYS[key]
        try:
            coerced = coercer(value)
        except (ValueError, TypeError) as e:
            return {"ok": False, "error": f"Invalid type for {key}: {e}"}
        if bounds is not None:
            lo, hi = bounds
            if not (lo <= coerced <= hi):
                return {
                    "ok": False,
                    "error": f"{key} must be between {lo} and {hi}, got {coerced}",
                }
        if key == "theme":
            coerced = normalize_theme_mode(coerced, DEFAULTS["theme"])
        if key == "snap_presses":
            coerced = normalize_snap_presses(coerced)
        validated[key] = coerced

    if self._draft is None:
        self._draft = {}
    self._draft.update(validated)

    return {"ok": True, "applied": validated}
```
New accent/density validation follows the existing `if key == "theme"` special-case pattern: add `if key == "accent"` and `if key == "density"` blocks that validate against allowed value sets.

**_strict_bool function** (new, module-level):
```python
# Place before SettingsState class definition, after imports

def _strict_bool(value):
    """Parse strict boolean values. Raises ValueError for ambiguous input."""
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

---

### `virelo/settings/persistence.py` (model, CRUD)

**Analog:** Self -- extending `__init__` and `save()` with new keys.

**QSettings read pattern -- boolean** (lines 15-18):
```python
self.enable_snap = _safe_bool(
    self._qs.value("enable_snap", DEFAULTS["enable_snap"], bool),
    DEFAULTS["enable_snap"],
)
```
New `minimize_to_tray` uses this exact pattern with `_safe_bool`.

**QSettings read pattern -- string** (lines 13-14):
```python
self.snap_key = str(self._qs.value("snap_key", DEFAULTS["snap_key"], str))
```
New `accent` and `density` use this pattern: `self.accent = str(self._qs.value("accent", DEFAULTS["accent"], str))`.

**QSettings write pattern** (lines 58-72):
```python
def save(self):
    self.clear()
    self._qs.beginGroup(SETTINGS_GROUP)
    self._qs.setValue("snap_key", self.snap_key)
    # ... one setValue per key ...
    self._qs.endGroup()
```
Add three new `setValue` calls for `accent`, `density`, `minimize_to_tray`.

**_safe_bool remains unchanged** (lines 84-97):
```python
def _safe_bool(val, default):
    if isinstance(val, bool):
        return val
    if val is None:
        return default
    text = str(val).strip().lower()
    if text in ("1", "true", "yes", "on"):
        return True
    if text in ("0", "false", "no", "off"):
        return False
    try:
        return bool(int(val))
    except Exception:
        return default
```
Per D-17, `_safe_bool` is permissive for QSettings reads. The strict parsing is in `_strict_bool` in `state.py`.

---

### `virelo/app/config.py` (config, static)

**Analog:** Self -- extending DEFAULTS dict.

**DEFAULTS dict pattern** (lines 17-29):
```python
DEFAULTS = {
    "snap_key": "shift",
    "restore_key": "ctrl",
    "enable_snap": True,
    "snap_presses": 3,
    "snap_interval": 1050,
    "width_pct": 76,
    "height_pct": 76,
    "ex_auto_size": False,
    "game_mode_enabled": True,
    "run_at_startup": False,
    "theme": "system",
}
```
Add three entries:
- `"accent": "slate"`
- `"density": "cozy"`
- `"minimize_to_tray": True`

---

### `virelo/app/window.py` (controller, event-driven)

**Analog:** Self -- refactoring existing key capture and tray menu logic.

**_on_capture_key -- current pattern (to be replaced)** (lines 290-304):
```python
def _on_capture_key(self, key: str):
    key_str = str(key).lower()
    if self._capture_target == "restore":
        self.settings.restore_key = key_str
        if hasattr(self, "_hotkey_listener"):
            self._hotkey_listener.update_restore_key(key_str)
        self._bridge.capture_status.emit("done")
        self._bridge.snap_status.emit(f"Restore key set to {key_str.upper()}.", 3000)
        self._bridge.settings_changed.emit(self._settings_state.get_json())
    else:
        if hasattr(self, "_hotkey_listener"):
            self._hotkey_listener.update_binding(key_str)
        self.key_captured.emit(key_str)
        self._bridge.capture_status.emit("done")
        self._bridge.settings_changed.emit(self._settings_state.get_json())
```
New pattern routes through draft instead of direct writes:
```python
# Replace with:
def _on_capture_key(self, key: str):
    key_str = str(key).lower()
    target_key = "restore_key" if self._capture_target == "restore" else "snap_key"
    self._settings_state.apply_draft({target_key: key_str})
    self._bridge.settings_changed.emit(self._settings_state.get_json())
    self._bridge.dirty_changed.emit(True)
    self._bridge.capture_status.emit("done")
    label = "Restore" if self._capture_target == "restore" else "Snap"
    self._bridge.snap_status.emit(f"{label} key set to {key_str.upper()}.", 3000)
```
Key difference: no direct `self.settings.X = value`, no `_hotkey_listener.update_*()` calls. Those happen in `_apply_side_effects` on commit.

**on_key_captured slot -- to be removed** (lines 259-262):
```python
@QtCore.Slot(str)
def on_key_captured(self, key: str):
    self.settings.snap_key = key
    self._bridge.snap_status.emit(f"Snap key set to {key.upper()}.", 3000)
```
Also remove: `self.key_captured.connect(self.on_key_captured)` at line 175, and the `key_captured = QtCore.Signal(str)` declaration at line 107.

**_toggle_run_at_startup -- current pattern (to be rerouted)** (lines 403-414):
```python
def _toggle_run_at_startup(self):
    try:
        if self.action_run_at_startup.isChecked():
            create_startup_shortcut()
            self.settings.run_at_startup = True
        else:
            remove_startup_shortcut()
            self.settings.run_at_startup = False
    except Exception as e:
        QtWidgets.QMessageBox.warning(self, "Error", f"Failed to modify startup shortcut:\n{e}")
        self.action_run_at_startup.setChecked(False)
    self.settings.save()
```
Tray handler should route through bridge: call `self._settings_state.apply_draft({"run_at_startup": checked})`, then `self._settings_state.commit_draft()`, then `self._bridge._apply_side_effects(applied)`, and emit signals.

**_apply_theme_mode -- side effect target** (lines 420-427):
```python
def _apply_theme_mode(self, mode: str):
    self._theme_mode = normalize_theme_mode(mode, DEFAULTS["theme"])
    self.settings.theme = self._theme_mode
    if self._theme_mode == "system":
        self._start_theme_sync()
    else:
        self._stop_theme_sync()
        self.set_theme(self._theme_mode)
```
This method should NOT write to `self.settings.theme` directly anymore (that is now handled by commit). Remove the `self.settings.theme = self._theme_mode` line. The method just handles visual application.

**Tray menu setup pattern** (lines 152-167):
```python
self.minimize_to_tray_on_exit = True
self.action_minimize_on_exit = menu.addAction("Minimize to Tray")
self.action_minimize_on_exit.setCheckable(True)
self.action_minimize_on_exit.setChecked(self.minimize_to_tray_on_exit)
self.action_minimize_on_exit.triggered.connect(self._toggle_minimize_on_exit)

self.action_run_at_startup = menu.addAction("Run at Startup")
self.action_run_at_startup.setCheckable(True)
self.action_run_at_startup.setChecked(bool(self.settings.run_at_startup))
self.action_run_at_startup.triggered.connect(self._toggle_run_at_startup)
```
`minimize_to_tray_on_exit` should be initialized from `self.settings.minimize_to_tray` instead of hardcoded `True`. Tray menu checked states should sync after each `commit_draft`.

---

### `frontend/src/app.jsx` (component, request-response)

**Analog:** Self -- extending existing patterns.

**bridgeToState mapping pattern** (lines 148-161):
```javascript
export function bridgeToState(settings) {
  return {
    snapEnabled: settings.enable_snap ?? true,
    snapKey: (settings.snap_key || 'shift').toUpperCase(),
    restoreKey: (settings.restore_key || 'ctrl').toUpperCase(),
    pressCount: settings.snap_presses ?? 3,
    interval: settings.snap_interval ?? 1050,
    width: settings.width_pct ?? 76,
    height: settings.height_pct ?? 76,
    gameMode: settings.game_mode_enabled ?? true,
    autoSize: settings.ex_auto_size ?? true,
    launchLogin: settings.run_at_startup ?? false,
  };
}
```
Add three new entries:
- `accent: settings.accent || 'slate'`
- `density: settings.density || 'cozy'`
- `minimizeToTray: settings.minimize_to_tray ?? true`

**stateToBridge mapping pattern** (lines 164-177):
```javascript
export function stateToBridge(state) {
  return JSON.stringify({
    enable_snap: state.snapEnabled,
    snap_key: state.snapKey.toLowerCase(),
    restore_key: state.restoreKey.toLowerCase(),
    snap_presses: state.pressCount,
    snap_interval: state.interval,
    width_pct: state.width,
    height_pct: state.height,
    game_mode_enabled: state.gameMode,
    ex_auto_size: state.autoSize,
    run_at_startup: state.launchLogin,
  });
}
```
Add three new entries:
- `accent: state.accent`
- `density: state.density`
- `minimize_to_tray: state.minimizeToTray`

**Signal subscription pattern** (lines 207-215):
```javascript
bridge.settings_changed.connect((json) => {
  try {
    const settings = JSON.parse(json);
    setState(bridgeToState(settings));
    setUnsaved(false);
  } catch (e) {
    console.error('[app] Failed to parse settings_changed:', e);
  }
});
```
New `dirty_changed` subscription follows the same `bridge.X.connect((val) => { ... })` pattern. Remove `setUnsaved(false)` from the `settings_changed` handler. Add:
```javascript
bridge.dirty_changed.connect((isDirty) => {
  setUnsaved(isDirty);
});
```

**set() function -- local setUnsaved to remove** (lines 229-237):
```javascript
const set = (p) => {
  setState((s) => {
    const next = { ...s, ...p };
    bridge.save_settings(stateToBridge(next), () => {});
    return next;
  });
  setUnsaved(true);
};
```
Remove `setUnsaved(true)` -- dirty state now comes from `dirty_changed` signal.

**handleSave/handleDiscard/handleReset -- setUnsaved(false) to remove** (lines 240-275):
- `handleSave` line 244: remove `if (r.ok) setUnsaved(false);`
- `handleDiscard` line 255: remove `if (r.ok) setUnsaved(false);`
- `handleReset` line 269: remove `setUnsaved(false);`

All dirty state is now driven by `dirty_changed` signal.

---

### `frontend/src/main.jsx` (component, event-driven)

**Analog:** Self -- refactoring initialization and theme handling.

**AppWithBridge tweaks initialization -- current pattern** (lines 12-19):
```javascript
function AppWithBridge({ bridge, initialTheme }) {
  const [tweaks, setTweaks] = React.useState({
    theme: initialTheme,
    accent: 'slate',
    density: 'cozy',
    radius: 6,
    sidebarMode: 'full',
  });
```
After this phase, `accent` and `density` should be initialized from bridge settings (loaded in Root), not hardcoded. `radius` and `sidebarMode` remain hardcoded per D-21.

**handleSetTweaks -- current pattern (to be refactored)** (lines 30-38):
```javascript
const handleSetTweaks = (updates) => {
  setTweaks((prev) => {
    const next = { ...prev, ...updates };
    // If theme changed, notify Python
    if (updates.theme && updates.theme !== prev.theme) {
      bridge.apply_theme(updates.theme, () => {});
    }
    return next;
  });
};
```
After this phase: remove `bridge.apply_theme` call (slot is removed). Theme, accent, density changes go through `app.set()` in pages.jsx, not through `handleSetTweaks`.

**Root get_theme_mode handling -- current pattern** (lines 57-73):
```javascript
React.useEffect(() => {
  getBridge().then((b) => {
    b.get_theme_mode((result) => {
      try {
        const r = JSON.parse(result);
        const mode = r.ok && r.data ? r.data : 'dark';
        setBridgeState({
          bridge: b,
          initialTheme: mode === 'light' ? 'light' : 'dark',
        });
      } catch (e) {
        console.error('[main] Failed to parse get_theme_mode result:', e);
        setBridgeState({ bridge: b, initialTheme: 'dark' });
      }
    });
  });
}, []);
```
After this phase, `get_theme_mode` returns `{mode, effective}` structure. Root should also call `get_settings` to load initial accent/density, and pass them to `AppWithBridge`.

---

### `frontend/src/pages.jsx` (component, event-driven)

**Analog:** `frontend/src/app.jsx` -- uses same `app.set()` pattern for all settings changes.

**GeneralPage accent control -- current pattern** (lines 169-179):
```javascript
<Row label="Accent color" description="Used for selection, toggles, and primary actions.">
  <div style={{ display: 'flex', gap: 6 }}>
    {Object.entries(ACCENTS).map(([k, v]) => (
      <button key={k} onClick={() => setTweaks({ accent: k })}
        style={{
          width: 22, height: 22, borderRadius: 11,
          background: t.isDark ? v.dark : v.light,
          border: tweaks.accent === k ? `2px solid ${t.text}` : `1px solid ${t.border}`,
          cursor: 'pointer', padding: 0,
        }}/>
    ))}
  </div>
</Row>
```
Change `setTweaks({ accent: k })` to `app.set({ accent: k })`. Read value from `app.accent` instead of `tweaks.accent`.

**GeneralPage density control** (lines 182-184):
```javascript
<Row label="Density" description="Controls spacing throughout the app." last>
  <Segmented options={[{ value: 'compact', label: 'Compact' }, { value: 'cozy', label: 'Cozy' }, { value: 'comfortable', label: 'Comfortable' }]}
    value={tweaks.density} onChange={(v) => setTweaks({ density: v })} />
</Row>
```
Change `setTweaks({ density: v })` to `app.set({ density: v })`. Read value from `app.density` instead of `tweaks.density`.

**GeneralPage theme control** (lines 165-167):
```javascript
<Row label="Theme" description="Light or dark surfaces throughout the app.">
  <Segmented options={[{ value: 'light', label: 'Light' }, { value: 'dark', label: 'Dark' }]}
    value={tweaks.theme} onChange={(v) => setTweaks({ theme: v })} />
</Row>
```
Change to include System option: `[{ value: 'system', label: 'System' }, { value: 'light', ... }, { value: 'dark', ... }]`. Route through `app.set({ themeMode: v })` instead of `setTweaks`. The Segmented value should track mode (system/light/dark), not effective theme.

**app.set() pattern from app.jsx** (line 229):
```javascript
const set = (p) => {
  setState((s) => {
    const next = { ...s, ...p };
    bridge.save_settings(stateToBridge(next), () => {});
    return next;
  });
};
```
This is the target pattern for accent/density/theme changes in pages.jsx. All go through `app.set({ key: value })` which calls `bridge.save_settings`.

---

### `frontend/src/bridge.js` (utility, static) -- reference only

**MOCK_BRIDGE and MOCK_SETTINGS** (lines 14-41):
Must be updated to:
- Add `accent`, `density`, `minimize_to_tray` to `MOCK_SETTINGS`
- Remove `apply_theme` and `toggle_run_at_startup` from `MOCK_BRIDGE`
- Update `get_theme_mode` to return `{mode, effective}` structure
- Add `dirty_changed: { connect: () => {} }` signal mock

---

## Shared Patterns

### Bridge Response Format
**Source:** `virelo/bridge/bridge.py` lines 66-73
**Apply to:** All bridge slot modifications and new slots
```python
# Success
return json.dumps({"ok": True, "data": result})
# Failure
return json.dumps({"ok": False, "error": str(e)})
```

### Signal Emission After State Change
**Source:** `virelo/bridge/bridge.py` lines 87-89, 101-104, 114-115
**Apply to:** All methods that modify draft state (save_settings, commit_draft, discard_draft, reset_defaults)
```python
# After any draft state change that succeeds:
self.settings_changed.emit(self._state.get_json())
self.dirty_changed.emit(self._state.has_draft)
```

### Side Effect If-Chain
**Source:** `virelo/bridge/bridge.py` lines 260-283
**Apply to:** New side effects for `run_at_startup`, `minimize_to_tray`
```python
if "key_name" in applied:
    # perform side effect
    # sync tray menu checkbox if applicable
```

### QSettings Read Pattern for New Keys
**Source:** `virelo/settings/persistence.py` lines 13-50
**Apply to:** New `accent`, `density`, `minimize_to_tray` keys
```python
# String key:
self.accent = str(self._qs.value("accent", DEFAULTS["accent"], str))
# Boolean key:
self.minimize_to_tray = _safe_bool(
    self._qs.value("minimize_to_tray", DEFAULTS["minimize_to_tray"], bool),
    DEFAULTS["minimize_to_tray"],
)
```

### QSettings Write Pattern
**Source:** `virelo/settings/persistence.py` lines 58-72
**Apply to:** New keys in `save()`
```python
self._qs.setValue("accent", self.accent)
self._qs.setValue("density", self.density)
self._qs.setValue("minimize_to_tray", self.minimize_to_tray)
```

### Frontend Signal Subscription
**Source:** `frontend/src/app.jsx` lines 206-215
**Apply to:** New `dirty_changed` subscription
```javascript
bridge.signal_name.connect((value) => {
  // update React state from Python signal
});
```

### Frontend Key Mapping
**Source:** `frontend/src/app.jsx` lines 148-177
**Apply to:** New accent, density, minimize_to_tray mappings
```javascript
// bridgeToState: python_snake_case -> reactCamelCase with fallback
accent: settings.accent || 'slate',
// stateToBridge: reactCamelCase -> python_snake_case
accent: state.accent,
```

## No Analog Found

No files in this phase lack analogs. Every modification extends existing patterns within the same file or uses a closely matching pattern from another file in the project.

## Metadata

**Analog search scope:** `virelo/bridge/`, `virelo/settings/`, `virelo/app/`, `frontend/src/`
**Files scanned:** 10
**Pattern extraction date:** 2026-04-24
