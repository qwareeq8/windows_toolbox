---
phase: 02-security-and-bridge-hardening
verified: 2026-04-24T19:45:00Z
status: verified
score: 5/5 must-haves verified
overrides_applied: 0
gaps_closed:
  - truth: "Every visible toggle, slider, button, and command palette action either calls a real bridge method or has been removed"
    resolution: "UI-04 redefined — Segmented preset controls (SHIFT/CTRL/ALT) are the intended design. Requirement text updated to match. User confirmed 2026-04-24."
  - truth: "Signal/slot consistency in reset_defaults"
    resolution: "Fixed in commit dff28e3 — line 129 now emits self._state.get_json() consistent with all other signal emissions."
human_verification:
  - test: "Launch the app and verify title bar minimize/close buttons work"
    expected: "Minimize button minimizes to tray or taskbar; close button closes the app"
    why_human: "Window management behavior requires running the app with a live QWebEngine host"
  - test: "Change a setting, click Save, close and reopen the app"
    expected: "The saved setting persists across restarts"
    why_human: "End-to-end persistence through QSettings requires running the app"
  - test: "Change a setting, click Discard, verify it reverts"
    expected: "The UI reverts to the previously saved value"
    why_human: "Draft/discard flow requires live bridge communication"
  - test: "Verify the WebEngine blocks navigation to https://example.com"
    expected: "Navigation is blocked, logged as a warning, page does not change"
    why_human: "Requires injecting a navigation request into the running WebEngine"
---

# Phase 2: Security and Bridge Hardening Verification Report

**Phase Goal:** The WebEngine host is locked down for an admin-elevated process, the bridge uses structured draft/commit state management, and every visible frontend control connects to the Python backend
**Verified:** 2026-04-24T19:45:00Z
**Status:** verified
**Re-verification:** Yes -- gaps closed 2026-04-24

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | Attempting to navigate the WebEngine to an external URL is blocked and does not load | VERIFIED | `webview.py` lines 101-118: `acceptNavigationRequest` blocks all schemes except `file://`, `data:`, and `http/https localhost` (dev mode only). Logs blocked attempts. |
| 2 | Changing a setting creates a draft; Save persists to QSettings and applies side effects; Discard reverts to persisted values | VERIFIED | `settings_state.py` has `apply_draft` (line 79), `commit_draft` (line 114), `discard_draft` (line 125). `bridge.py` `save_settings` calls `apply_draft` (line 84) NOT `_apply_side_effects`. `commit_draft` slot (line 96) calls `_apply_side_effects` (line 102). `discard_draft` slot (line 108) clears draft and emits signal. Frontend `handleSave` calls `bridge.commit_draft` (app.jsx line 241), `handleDiscard` calls `bridge.discard_draft` (app.jsx line 253). |
| 3 | No fake controls are visible in the UI | VERIFIED | grep for `rememberCols`, `showHidden`, `showExts`, `startTray`, `autoUpdate`, `telemetry` across `frontend/src/` returns zero matches. grep for `Hidden files`, `Remember column`, `Automatic updates`, `Anonymous telemetry`, `Check for updates`, `up to date` returns zero matches. ExplorerPage has only one Card with auto-size toggle. ShortcutsPage has exactly 3 items. AboutPage shows icon, name, version, license only. |
| 4 | Every visible control calls a real bridge method or has been removed | VERIFIED | All command palette actions wired (onTestSnap, onSave, onReset in panels.jsx lines 25-27). Title bar minimize/close call `bridge.setWindowCommand` (app.jsx lines 35,42). All toggles call `app.set()` which calls `bridge.save_settings`. UI-04 redefined: Segmented preset controls are the intended design (user confirmed 2026-04-24). |
| 5 | The bridge returns structured payloads for all operations and unknown keys produce structured error | VERIFIED | All 13 bridge slots return `{"ok": true/false, ...}` payloads (37 occurrences of `"ok"` in bridge.py). `settings_state.py` `apply_draft` rejects unknown keys with `{"ok": False, "error": "Unknown keys: [...]"}` (line 88-89). `get_settings`, `get_theme_mode`, `get_launch_at_login`, `get_snap_enabled` all return `{"ok": true, "data": ...}` structure. |

**Score:** 4/5 truths verified

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `webview.py` | WebEngine security hardening | VERIFIED | Contains `acceptNavigationRequest` override (line 101), `_MISSING_FRONTEND_HTML` (line 32), `_is_dev_mode` checks only env var (line 76), `NoContextMenu` policy (line 164), `LocalContentCanAccessRemoteUrls` conditioned on `_is_dev_mode()` (line 150) |
| `settings_state.py` | Draft state model | VERIFIED | Contains `_draft = None` init (line 46), `apply_draft` (line 79), `commit_draft` (line 114), `discard_draft` (line 125), `has_draft` property (line 74-77), `get_all()` overlays draft (line 60-61). `apply_draft` does NOT call `self._settings.save()`. `commit_draft` DOES call `self._settings.save()` (line 120). |
| `bridge.py` | Structured payloads, draft slots, window commands | VERIFIED | Contains `commit_draft` slot (line 96), `discard_draft` slot (line 108), `has_draft` slot (line 119), `setWindowCommand` slot (line 239). All slots return structured JSON. `save_settings` calls `apply_draft` not `apply_partial`. No `apply_partial` references remain. |
| `frontend/src/pages.jsx` | Cleaned settings pages with no fake controls | VERIFIED | ExplorerPage has one Card with one Row (auto-size, line 115-124). ShortcutsPage has 3 items (lines 129-133). GeneralPage has no fake rows. AboutPage has icon, name, version, license only (lines 203-230). |
| `frontend/src/app.jsx` | Draft-aware save/discard, title bar wiring, cleaned state mappings | VERIFIED | `handleSave` calls `bridge.commit_draft` (line 241). `handleDiscard` calls `bridge.discard_draft` (line 253). `handleReset` unwraps `r.data` (line 267). Initial `get_settings` unwraps `r.data` (line 198). TitleBar has minimize/close buttons calling `setWindowCommand` (lines 35,42). No maximize button. `bridgeToState` has no fake keys. `stateToBridge` has no `theme: undefined`. Sidebar shows `v{version}` without "up to date". |
| `frontend/src/panels.jsx` | Command palette with real action handlers | VERIFIED | `CommandPalette` accepts `onTestSnap`, `onSave`, `onReset` props (line 8). "Test snap" calls `onTestSnap?.()` (line 25). "Save changes" calls `onSave?.()` (line 26). "Reset to defaults" calls `onReset?.()` (line 27). No `() => {}` no-ops. No `_saved` references. |
| `frontend/src/bridge.js` | Mock bridge with new slots and structured responses | VERIFIED | MOCK_BRIDGE has `commit_draft` (line 24), `discard_draft` (line 25), `has_draft` (line 26), `setWindowCommand` (line 35). `get_settings` returns `{ok: true, data: MOCK_SETTINGS}` (line 22). All other slots return structured payloads. |
| `frontend/src/main.jsx` | Structured response handling for get_theme_mode | VERIFIED | `get_theme_mode` callback parses JSON and unwraps `r.data` (lines 61-63). Falls back to 'dark' on parse failure (line 69). |

### Key Link Verification

| From | To | Via | Status | Details |
|------|----|-----|--------|---------|
| `bridge.py:save_settings` | `settings_state.py:apply_draft` | `self._state.apply_draft(data)` | WIRED | Line 84 of bridge.py calls apply_draft, NOT apply_partial |
| `bridge.py:commit_draft` | `settings_state.py:commit_draft` | `self._state.commit_draft()` | WIRED | Line 99 of bridge.py |
| `bridge.py:commit_draft` | `bridge.py:_apply_side_effects` | Side effects on commit only | WIRED | Line 102: `self._apply_side_effects(result.get("applied", {}))` only in commit_draft |
| `app.jsx:handleSave` | `bridge.commit_draft` | Bridge slot call | WIRED | Line 241: `bridge.commit_draft((result) => {...})` |
| `app.jsx:handleDiscard` | `bridge.discard_draft` | Bridge slot call | WIRED | Line 253: `bridge.discard_draft((result) => {...})` |
| `app.jsx:bridgeToState` | `bridge.get_settings` response | Unwrapping `result.data` | WIRED | Lines 197-200: `r.ok && r.data` -> `bridgeToState(r.data)` |
| `app.jsx:TitleBar` | `bridge.setWindowCommand` | Button onClick | WIRED | Lines 35,42: minimize and close buttons call `bridge.setWindowCommand` |
| `panels.jsx:CommandPalette` | `app.jsx:handleTestSnap/Save/Reset` | Props | WIRED | Line 329-330: props passed from VireloApp, received in panels.jsx line 8 |
| `webview.py:VireloWebPage` | `QWebEnginePage` | acceptNavigationRequest override | WIRED | Class inherits QWebEnginePage (line 98), overrides method (line 101) |

### Data-Flow Trace (Level 4)

| Artifact | Data Variable | Source | Produces Real Data | Status |
|----------|---------------|--------|--------------------|--------|
| `app.jsx` | `state` (via `useState`) | `bridge.get_settings` -> `r.data` -> `bridgeToState()` | Yes -- Python `SettingsState.get_all()` reads from QSettings | FLOWING |
| `settings_state.py` | `_draft` / `get_all()` | `self._settings` attributes from QSettings | Yes -- `getattr(self._settings, key, ...)` reads persisted values | FLOWING |

### Behavioral Spot-Checks

Step 7b: SKIPPED (no runnable entry points -- app requires PySide6/QWebEngine which cannot be tested headlessly in this environment)

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|------------|-------------|--------|----------|
| SEC-01 | 02-01 | WebEngine blocks external navigation | SATISFIED | `acceptNavigationRequest` blocks all non-local URLs |
| SEC-02 | 02-01 | LocalContentCanAccessRemoteUrls disabled in release | SATISFIED | Line 150: set to `_is_dev_mode()` |
| SEC-03 | 02-01 | Dev mode triggered only by VIRELO_DEV=1 | SATISFIED | `_is_dev_mode()` checks only env var, no `sys.frozen` fallback |
| SEC-04 | 02-01 | Missing frontend shows error message, not blank page | SATISFIED | `_MISSING_FRONTEND_HTML` rendered via `setHtml()` when URL is None |
| SEC-05 | 02-01 | Context menu disabled in release mode | SATISFIED | `NoContextMenu` policy set when `not _is_dev_mode()` |
| BRDG-01 | 02-02 | Draft model holds unsaved changes separately | SATISFIED | `_draft` dict in SettingsState, `apply_draft` stores without persisting |
| BRDG-02 | 02-02 | Save commits draft and applies side effects | SATISFIED | `commit_draft` persists to QSettings, calls `_apply_side_effects` |
| BRDG-03 | 02-02 | Discard reverts to persisted settings | SATISFIED | `discard_draft` clears `_draft`, emits signal with persisted values |
| BRDG-04 | 02-02 | Unknown keys rejected with structured error | SATISFIED | `apply_draft` line 88-89: unknown keys produce `{"ok": false, "error": ...}` |
| BRDG-05 | 02-02 | Structured payloads for all operations | SATISFIED | All 13 slots return `{"ok": true/false, ...}` JSON |
| BRDG-06 | 02-02 | Startup shortcut on save only | SATISFIED | Side effects fire only in `commit_draft`, not `save_settings` |
| UI-01 | 02-03 | Fake controls removed | SATISFIED | Zero grep matches for any fake control text or state keys |
| UI-02 | 02-03 | No-op command palette actions wired | SATISFIED | onTestSnap, onSave, onReset all wired to real handlers |
| UI-03 | 02-03 | Title bar minimize/close work | SATISFIED | Buttons call `bridge.setWindowCommand` with minimize/close |
| UI-04 | 02-03 | Snap/restore key selection via presets | SATISFIED | Segmented preset controls (SHIFT/CTRL/ALT) are the intended design. Requirement redefined to match implementation. User confirmed 2026-04-24. |
| UI-05 | 02-03 | Every visible control connects to bridge | SATISFIED | All toggles, sliders, buttons call app.set() -> bridge.save_settings, or direct bridge methods |
| UI-06 | 02-03 | React state survives save/discard/reset/push | SATISFIED | handleSave/handleDiscard/handleReset all update state correctly; settings_changed signal handler updates state from pushed values |

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
|------|------|---------|----------|--------|
| *(none remaining)* | | All anti-patterns resolved | | Fixed in commit dff28e3 |

### Human Verification Required

### 1. Title bar window controls

**Test:** Launch Virelo, click the minimize button (dash), then click the close button (x)
**Expected:** Minimize sends the window to taskbar/tray; close exits or hides to tray per settings
**Why human:** Window management behavior requires a running PySide6 application

### 2. Draft save/discard persistence

**Test:** Change a setting (e.g., width slider), click Save, close and reopen the app
**Expected:** The saved setting persists across restarts; clicking Discard instead reverts
**Why human:** End-to-end QSettings persistence requires running the full app stack

### 3. WebEngine navigation blocking

**Test:** In dev tools (or via injected JS), attempt `window.location = 'https://example.com'`
**Expected:** Navigation blocked, warning logged, page unchanged
**Why human:** Requires live WebEngine to test navigation filter

### 4. Missing frontend error page

**Test:** Delete `frontend/dist/index.html`, launch the app without VIRELO_DEV=1
**Expected:** Styled error page appears with "Frontend build not found" and build instructions
**Why human:** Requires launching the app without a frontend build

### Gaps Summary

**0 gaps remaining** (2 closed):

**Gap 1 -- UI-04 key capture (CLOSED):** Requirement redefined. Segmented preset controls (SHIFT/CTRL/ALT) are the intended design. User confirmed 2026-04-24. REQUIREMENTS.md updated.

**Gap 2 -- reset_defaults signal emission (CLOSED):** Fixed in commit `dff28e3`. Line 129 now emits `self._state.get_json()` consistent with all other signal emissions.

---

_Verified: 2026-04-24T19:45:00Z_
_Verifier: Claude (gsd-verifier)_
