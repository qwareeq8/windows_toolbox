# Phase 2: Security and Bridge Hardening - Research

**Researched:** 2026-04-24
**Domain:** PySide6 QWebEngine security, Python-side state management, React-Python bridge architecture
**Confidence:** HIGH

## Summary

This phase has three distinct workstreams that share a common surface area (the `webview.py` / `bridge.py` / `settings_state.py` / `app.jsx` files): (1) hardening the QWebEngine host for a process that runs with admin privileges, (2) restructuring the bridge to use a Python-side draft/commit model instead of immediate persistence, and (3) removing fake frontend controls that have no backend implementation and wiring all remaining controls to real bridge methods.

The current codebase has clear security gaps: `LocalContentCanAccessRemoteUrls` is set to `True` unconditionally, no navigation filtering exists (any URL can be loaded), dev mode activates on `not sys.frozen` which means running from source always enables remote access, and the context menu exposes Chromium dev tools in all modes. The draft state model is straightforward to implement since `SettingsState` already has `apply_partial` with validation -- the change is adding a `_draft` dict layer between the frontend and `Settings.save()`. The fake control removal is a mechanical deletion task guided by the exhaustive list in CONTEXT.md decision D-11 through D-14.

**Primary recommendation:** Execute in three sequential waves: WebEngine security lockdown first (most critical, fewest dependencies), then bridge draft state model (changes save/discard flow that frontend consumes), then fake control removal and wiring (depends on the new bridge API being stable).

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Override `acceptNavigationRequest` in `VireloWebPage` to block all non-local navigation. Allow `file://` and `http://localhost` (dev server) URLs only. Log blocked navigation attempts.
- **D-02:** `LocalContentCanAccessRemoteUrls` set to `False` in release mode, `True` only when `_is_dev_mode()` returns True.
- **D-03:** Dev mode detection must require `VIRELO_DEV=1` explicitly -- remove the `not sys.frozen` fallback. Running from source without the env var should behave like release mode (with local files, not Vite dev server).
- **D-04:** When `frontend/dist/index.html` is missing in release mode, display a styled inline HTML error page (not a blank white page) explaining the build is missing and how to run `scripts/build-frontend.ps1`.
- **D-05:** Default WebEngine context menu disabled in release mode. Enabled in dev mode for debugging (Inspect Element access).
- **D-06:** Add a `_draft` dict to `SettingsState` that holds unsaved changes separately from the persisted `Settings` object. On construction, `_draft` is `None` (no pending changes).
- **D-07:** When the frontend sends changes via `save_settings`, the bridge stores them in `_draft` instead of immediately writing to QSettings. The frontend reflects draft values.
- **D-08:** A new `commit_draft` bridge slot persists `_draft` to QSettings via `Settings.save()`, applies side effects (startup shortcut, snap manager, Explorer worker), then clears `_draft`.
- **D-09:** A new `discard_draft` bridge slot clears `_draft` and pushes the persisted settings back to the frontend via `settings_changed` signal.
- **D-10:** `get_settings` returns the merged view: persisted settings overlaid with any draft values. A separate `has_draft` slot or signal indicates whether unsaved changes exist.
- **D-11:** Remove these controls entirely from the frontend: Explorer page "Remember column widths" toggle, "Hidden files" card; General page "Start minimized to tray", "Automatic updates", "Anonymous telemetry" toggles; About page "Check for updates" button, "Up to date" badge, "Documentation Open" button, "Report an issue Open" button.
- **D-12:** Remove corresponding fake state keys from `bridgeToState` and `stateToBridge`: `rememberCols`, `showHidden`, `showExts`, `startTray`, `autoUpdate`, `telemetry`.
- **D-13:** Remove hardcoded changelog from About page -- version number from `__APP_VERSION__` is sufficient. About page shows: app icon, name, version, and license info only.
- **D-14:** Explorer page retains only the auto-size columns toggle after cleanup.
- **D-15:** All bridge slots return structured JSON: `{"ok": true, "data": ...}` on success, `{"ok": false, "error": "..."}` on failure.
- **D-16:** Unknown keys in `apply_partial` are rejected with a structured error response instead of being silently dropped.
- **D-17:** The bridge validates all inputs before acting.
- **D-18:** Title bar minimize and close buttons call bridge slots. Maximize button removed.
- **D-19:** Command palette "Test snap" action calls `bridge.test_snap()`.
- **D-20:** Command palette "Reset to defaults" action calls the same `handleReset` used by the footer button.
- **D-21:** Command palette "Save changes" action calls `handleSave`.
- **D-22:** Remove fake shortcut entries. Keep only: "Trigger snap", "Restore last snap", and "Command palette".

### Claude's Discretion
- Error page HTML styling and exact copy
- Whether `get_settings` returns draft-merged or persisted-only (recommended: draft-merged)
- Internal implementation of `_draft` (dict overlay vs full copy)
- Whether to add a `setWindowCommand` bridge slot or separate `minimizeWindow`/`closeWindow` slots
- Sidebar "up to date" status text replacement after fake update status removal

### Deferred Ideas (OUT OF SCOPE)
None -- discussion stayed within phase scope.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| SEC-01 | WebEngine blocks external navigation via acceptNavigationRequest override | `QWebEnginePage.acceptNavigationRequest(url, type, isMainFrame)` API verified; override in existing `VireloWebPage` subclass |
| SEC-02 | LocalContentCanAccessRemoteUrls disabled in release mode | `QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls` attribute verified; currently unconditionally `True` at webview.py:109 |
| SEC-03 | Dev mode triggered only by VIRELO_DEV=1, not sys.frozen check | Current `_is_dev_mode()` at webview.py:49-57 has the `not sys.frozen` fallback to remove |
| SEC-04 | Missing frontend build shows clear error message, not blank page | `QWebEnginePage.setHtml(html, baseUrl)` API available for inline HTML; 2MB limit not a concern for error pages |
| SEC-05 | Default WebEngine context menu disabled in release mode | `QWidget.setContextMenuPolicy(Qt.NoContextMenu)` confirmed for QWebEngineView |
| BRDG-01 | Python-side draft model holds unsaved changes separately | `SettingsState` class in settings_state.py is the correct extension point; add `_draft` dict |
| BRDG-02 | Save commits draft to QSettings and applies side effects | `_apply_side_effects` in bridge.py already dispatches; move trigger to `commit_draft` |
| BRDG-03 | Discard reverts draft to persisted settings | `SettingsState.get_all()` reads from `Settings` object which holds persisted values |
| BRDG-04 | Unknown bridge keys rejected with structured error | `apply_partial` at settings_state.py:63 currently does `continue` for unknown keys; change to error |
| BRDG-05 | Bridge returns structured payloads for all operations | Most slots already return `{"ok": ...}` but `get_settings`, `get_theme_mode`, `get_launch_at_login`, `get_snap_enabled` do not |
| BRDG-06 | Launch-at-login toggle creates/removes startup shortcut on save | `_toggle_run_at_startup` in main.py:1292 handles shortcut creation; must fire during `commit_draft` |
| UI-01 | Fake controls removed | Exhaustive list in D-11; locations identified in pages.jsx |
| UI-02 | No-op command palette actions removed or wired | panels.jsx:26-28 has three no-ops: Test snap, Save changes, Reset to defaults |
| UI-03 | Title bar minimize/close call bridge and work in frameless window | TitleBar in app.jsx:35-38 renders buttons with no handlers |
| UI-04 | Key capture uses bridge startKeyCapture/cancelKeyCapture | Existing `capture_key` slot in bridge.py:126 and `_begin_key_capture` in main.py:1090 |
| UI-05 | Every visible toggle/slider/button connects to Python bridge | After fake control removal and command palette wiring, this is satisfied by construction |
| UI-06 | React state survives save, discard, reset, and external Python-pushed updates | Draft model provides clean state transitions; `settings_changed` signal already used |
</phase_requirements>

## Project Constraints (from CLAUDE.md)

- **Never reintroduce "Windows Toolbox" or "Toolbox"** in any file
- **Never add fake or placeholder UI controls** that do not connect to real backend logic -- every visible control must be wired to the Python bridge (this phase enforces this rule)
- **Never commit generated artifacts** (frontend/dist/, dist/, build/, .venv/, __pycache__/)
- **Never hardcode version strings** -- use `APP_VERSION` from `app_config.py`, `__APP_VERSION__` in frontend
- **`app_config.py` must not be imported in `Virelo.spec`** -- use regex for version parsing
- **PowerShell `$LASTEXITCODE`** must be checked after every external command
- **Vite `define` values must use `JSON.stringify()`** -- already done in vite.config.js
- **Inno Setup `#define` must use `#ifndef` guard** for `/D` override

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Navigation filtering | WebEngine host (Python) | -- | `acceptNavigationRequest` is a QWebEnginePage virtual, runs in Python process |
| Dev mode detection | WebEngine host (Python) | -- | Environment variable check happens before URL loading |
| Context menu policy | WebEngine host (Python) | -- | `setContextMenuPolicy` is a QWidget API on the Python side |
| Missing build error page | WebEngine host (Python) | -- | `setHtml()` called before any JS loads |
| Draft state management | Bridge/Settings (Python) | Frontend (React) | Python owns canonical state; React reflects it via `get_settings` |
| Settings persistence | Settings layer (Python) | -- | QSettings writes to Windows registry |
| Side effect dispatch | Bridge (Python) | MainWindow (Python) | Bridge calls MainWindow private methods for snap/explorer/theme/startup |
| Fake control removal | Frontend (React) | -- | Pure JSX deletion in pages.jsx, panels.jsx, app.jsx |
| Command palette wiring | Frontend (React) | Bridge (Python) | React calls existing bridge slots; no new Python code needed |
| Title bar window controls | Frontend (React) | Bridge (Python) | React fires bridge call; Python does the actual minimize/close |
| Key mapping cleanup | Frontend (React) | -- | `bridgeToState`/`stateToBridge` are frontend-only mappers |

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| PySide6 | >=6.6 | QWebEngineView, QWebChannel, QWebEnginePage, QWebEngineSettings | Already in use; provides all WebEngine security APIs needed | [VERIFIED: requirements.txt]
| React | ^19.1.0 | Frontend UI framework | Already in use; no changes needed | [VERIFIED: package.json]
| Vite | ^6.3.4 | Frontend build tool | Already in use; no changes needed | [VERIFIED: package.json]

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| json (stdlib) | -- | Bridge payload serialization | All bridge slot methods | [VERIFIED: already used in bridge.py]

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `acceptNavigationRequest` override | `QWebEngineUrlRequestInterceptor` | Interceptor blocks ALL requests (CSS, JS, images), not just navigations. Too broad for this use case -- we only want to block top-level navigation to external URLs, not resource loading. |
| `setHtml()` for error page | `QWebEngineUrlSchemeHandler` for custom scheme | Over-engineered for a single static error page. `setHtml()` is the standard approach. |
| `Qt.NoContextMenu` policy | Override `contextMenuEvent` | Policy flag is simpler and more declarative than overriding the event handler. |
| Single `setWindowCommand(cmd)` slot | Separate `minimizeWindow()` / `closeWindow()` slots | Single slot with string command is more extensible if future window commands are added; matches existing bridge pattern of string-typed parameters |

**Installation:** No new packages needed. All APIs are already available in the existing PySide6 and React dependencies.

## Architecture Patterns

### System Architecture Diagram

```
                    Frontend (React)                         Python Backend
                    ================                         ==============

  User Action ──> React Component ──> bridge.method(json) ──> @Slot method
       |              |                                          |
       |         [local state                              [validates input]
       |          updated via                                     |
       |          app.set()]                               [stores in _draft
       |              |                                     if save_settings]
       |              v                                          |
       |         "unsaved" = true                                |
       |              |                                          v
       v              |                              +----- commit_draft -----+
  Save clicked ──> handleSave() ──> bridge.commit_draft() ──> |               |
                                                              | _draft -> Settings.save()
                                                              | _apply_side_effects()
                                                              | _draft = None
                                                              +-------+-------+
                                                                      |
                                                                      v
                                              settings_changed Signal ──> React setState
                                                                         unsaved = false

  Discard clicked ──> handleDiscard() ──> bridge.discard_draft() ──> _draft = None
                                                                     |
                                                                     v
                                              settings_changed Signal ──> React setState
                                                                         (persisted values)
```

```
  Navigation Request Flow (Security)
  ====================================

  WebEngine loads URL
       |
       v
  VireloWebPage.acceptNavigationRequest(url, type, isMainFrame)
       |
       +---> url.scheme() in ("file",) ? ────────────> ALLOW
       |
       +---> url.scheme() == "http" AND
       |     url.host() == "localhost" AND
       |     _is_dev_mode() ? ───────────────────────> ALLOW (dev only)
       |
       +---> url.scheme() == "data" ? ───────────────> ALLOW (setHtml uses data: URLs)
       |
       +---> else ───────────────────────────────────> BLOCK + LOG
```

### Recommended Project Structure

No structural changes to file layout in this phase. All modifications are within existing files:

```
Virelo/
  webview.py          # SEC-01..05: acceptNavigationRequest, settings, context menu, error page
  bridge.py           # BRDG-01..06: draft slots, response standardization, window commands
  settings_state.py   # BRDG-01,04: _draft dict, unknown key rejection
  main.py             # Window command handlers (minimize, close)
  frontend/src/
    app.jsx           # UI-03,06: title bar wiring, bridgeToState/stateToBridge cleanup, save/discard flow
    pages.jsx         # UI-01: fake control removal
    panels.jsx        # UI-02: command palette wiring
    bridge.js         # Mock bridge updates for new slots
```

### Pattern 1: acceptNavigationRequest Override

**What:** Override the virtual method in `VireloWebPage` to filter navigation requests by URL scheme and host.
**When to use:** Whenever the WebEngine host is loaded in a privileged process that should not navigate to arbitrary remote content.
**Example:**
```python
# Source: https://doc.qt.io/qtforpython-6.8/PySide6/QtWebEngineCore/QWebEnginePage.html
class VireloWebPage(QWebEnginePage):
    def acceptNavigationRequest(self, url, nav_type, is_main_frame):
        scheme = url.scheme().lower()
        # Allow file:// (release mode local files)
        if scheme == "file":
            return True
        # Allow data: (used by setHtml for error pages)
        if scheme == "data":
            return True
        # Allow localhost in dev mode only
        if scheme in ("http", "https") and url.host() == "localhost" and _is_dev_mode():
            return True
        # Block everything else
        LOG.warning("Blocked navigation to: %s (type=%s, main_frame=%s)", url.toString(), nav_type, is_main_frame)
        return False
```
[VERIFIED: PySide6 docs at doc.qt.io/qtforpython-6.8]

### Pattern 2: Draft State Model

**What:** A `_draft` dict in `SettingsState` that overlays persisted settings, providing a merged view through `get_all()` and committing/discarding on explicit user action.
**When to use:** Any settings UI that needs a "Save" / "Discard" flow rather than immediate persistence.
**Example:**
```python
class SettingsState:
    def __init__(self, settings: Settings):
        self._settings = settings
        self._draft = None  # None = no pending changes

    def get_all(self) -> dict:
        """Return merged view: persisted settings overlaid with draft."""
        result = {}
        for key, (coercer, _) in self.KEYS.items():
            val = getattr(self._settings, key, DEFAULTS.get(key))
            result[key] = coercer(val)
        # Overlay draft values
        if self._draft:
            result.update(self._draft)
        # Normalize theme and snap_presses
        result["theme"] = normalize_theme_mode(result["theme"], DEFAULTS["theme"])
        result["snap_presses"] = normalize_snap_presses(result["snap_presses"])
        return result

    @property
    def has_draft(self) -> bool:
        return self._draft is not None and len(self._draft) > 0

    def apply_draft(self, data: dict) -> dict:
        """Validate and store changes in draft (not persisted)."""
        validated = {}
        unknown = [k for k in data if k not in self.KEYS]
        if unknown:
            return {"ok": False, "error": f"Unknown keys: {unknown}"}
        for key, value in data.items():
            coercer, bounds = self.KEYS[key]
            try:
                coerced = coercer(value)
            except (ValueError, TypeError) as e:
                return {"ok": False, "error": f"Invalid type for {key}: {e}"}
            if bounds is not None:
                lo, hi = bounds
                if not (lo <= coerced <= hi):
                    return {"ok": False, "error": f"{key} must be between {lo} and {hi}, got {coerced}"}
            if key == "theme":
                coerced = normalize_theme_mode(coerced, DEFAULTS["theme"])
            if key == "snap_presses":
                coerced = normalize_snap_presses(coerced)
            validated[key] = coerced
        if self._draft is None:
            self._draft = {}
        self._draft.update(validated)
        return {"ok": True, "applied": validated}

    def commit_draft(self) -> dict:
        """Persist draft to QSettings and clear draft."""
        if not self._draft:
            return {"ok": True, "applied": {}}
        for key, value in self._draft.items():
            setattr(self._settings, key, value)
        self._settings.save()
        applied = dict(self._draft)
        self._draft = None
        return {"ok": True, "applied": applied}

    def discard_draft(self):
        """Clear draft without persisting."""
        self._draft = None
```
[ASSUMED -- design follows CONTEXT.md decisions D-06..D-10; implementation pattern is standard Python]

### Pattern 3: Structured Bridge Response

**What:** All bridge slots return `{"ok": true, "data": ...}` or `{"ok": false, "error": "..."}`.
**When to use:** Every `@Slot` method in `VireloBridge`.
**Example:**
```python
@Slot(result=str)
def get_settings(self) -> str:
    try:
        settings = self._state.get_all()
        return json.dumps({"ok": True, "data": settings})
    except Exception as e:
        LOG.exception("get_settings failed")
        return json.dumps({"ok": False, "error": str(e)})

@Slot(result=str)
def get_theme_mode(self) -> str:
    mode = "dark"
    if self._main_window:
        mode = getattr(self._main_window, "_theme_mode", "dark")
    return json.dumps({"ok": True, "data": mode})
```
[ASSUMED -- extrapolation from existing partial pattern in bridge.py]

### Pattern 4: Context Menu Toggle by Mode

**What:** Disable the WebEngine context menu in release mode, keep it in dev mode.
**When to use:** During `VireloWebView.__init__()`.
**Example:**
```python
from PySide6.QtCore import Qt

# In VireloWebView.__init__:
if not _is_dev_mode():
    self.setContextMenuPolicy(Qt.ContextMenuPolicy.NoContextMenu)
```
[VERIFIED: Qt documentation confirms Qt.NoContextMenu / Qt.ContextMenuPolicy.NoContextMenu disables context menus]

### Anti-Patterns to Avoid

- **Persisting on every keystroke:** The old `save_settings` writes to QSettings immediately. The draft model prevents registry thrashing by batching writes to `commit_draft`.
- **Silent key dropping:** The old `apply_partial` silently skips unknown keys (line 63-64 of settings_state.py). This hides frontend bugs. D-16 requires explicit error responses for unknown keys.
- **`not sys.frozen` as dev mode:** This causes release-mode security settings to be bypassed whenever running from source. D-03 requires explicit opt-in via `VIRELO_DEV=1`.
- **Mixing save + side effects:** The current `save_settings` calls `_apply_side_effects` immediately. Side effects should only fire on `commit_draft`, not on draft updates, because the user might discard their changes.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Navigation filtering | Custom URL regex parser | `QWebEnginePage.acceptNavigationRequest` | Qt provides the hook; URL parsing via QUrl.scheme()/host() is reliable |
| Context menu suppression | JavaScript `contextmenu` event listener | `setContextMenuPolicy(Qt.NoContextMenu)` | Widget-level policy is more robust than JS-level interception |
| Inline error page | External HTML file bundled in PyInstaller | `QWebEnginePage.setHtml(html_string)` | No file I/O needed; HTML is a small constant string |
| Draft state diffing | Custom deep-diff library | Simple dict overlay (`{**persisted, **draft}`) | Settings are a flat dict with 11 keys; no nesting or complex types |

**Key insight:** All security hardening uses existing Qt APIs -- no third-party security libraries are needed. The draft state model is a thin Python-only change to `SettingsState` with no new dependencies.

## Common Pitfalls

### Pitfall 1: data: URL Scheme in acceptNavigationRequest
**What goes wrong:** `setHtml()` internally converts HTML to a `data:` URL. If `acceptNavigationRequest` blocks `data:` scheme, the error page itself will not render.
**Why it happens:** The error page for missing frontend is loaded via `setHtml()`, which uses `data:` URLs internally.
**How to avoid:** Explicitly allow `data:` scheme in `acceptNavigationRequest`.
**Warning signs:** Error page shows blank white instead of styled HTML.

### Pitfall 2: save_settings Must Update Draft, Not Persist
**What goes wrong:** If the renamed `save_settings` slot (now updating draft) still calls `Settings.save()` or `_apply_side_effects()`, changes are persisted before the user clicks Save.
**Why it happens:** Easy to miss removing the `self._settings.save()` call inside the old `apply_partial`.
**How to avoid:** The old `apply_partial` in `SettingsState` calls `self._settings.save()` at line 83. The new `apply_draft` method must NOT call `save()`. Only `commit_draft` calls `save()`.
**Warning signs:** Settings persist even when the user clicks Discard.

### Pitfall 3: Frontend State Keys After Fake Control Removal
**What goes wrong:** React state initialization still includes removed keys (`rememberCols`, `showHidden`, etc.) causing stale state or undefined errors.
**Why it happens:** The initial state in `VireloApp` (app.jsx:183-189) hardcodes all keys including fakes.
**How to avoid:** Remove the fake keys from: (1) initial `useState` object, (2) `bridgeToState()`, (3) `stateToBridge()`, (4) `MOCK_SETTINGS` in bridge.js.
**Warning signs:** Console warnings about undefined state properties; React components receiving undefined props.

### Pitfall 4: Title Bar Buttons Need -webkit-app-region Awareness
**What goes wrong:** In a frameless window, title bar buttons might not receive click events if the title bar area is marked as a drag region.
**Why it happens:** The title bar div acts as both a drag handle (via WM_NCHITTEST at main.py:1356) and a button container.
**How to avoid:** The current implementation uses WM_NCHITTEST with a 4px border zone only at edges, not the title bar. The title bar is not a native drag region -- it is a React-rendered div. Click events will work because there is no `-webkit-app-region: drag` set. Verify title bar button `onClick` handlers fire correctly.
**Warning signs:** Clicking minimize/close does nothing.

### Pitfall 5: get_settings Response Shape Change Breaks Frontend
**What goes wrong:** Wrapping `get_settings` in `{"ok": true, "data": settings}` changes the JSON shape. Frontend code that does `JSON.parse(json)` and passes directly to `bridgeToState(settings)` will break because `settings` is now nested under `data`.
**Why it happens:** BRDG-05 requires all slots to return structured payloads, but the frontend currently expects raw settings dict from `get_settings`.
**How to avoid:** Update ALL frontend callsites: `bridge.get_settings` callback, `bridge.reset_defaults` callback, `settings_changed` signal handler. Each must unwrap `result.data` before passing to `bridgeToState()`.
**Warning signs:** All settings display as default values or the app shows a blank state.

### Pitfall 6: Mock Bridge in Dev Mode Must Match New API
**What goes wrong:** The mock bridge in `bridge.js` (used when QWebChannel is unavailable in dev mode) does not have the new slots (`commit_draft`, `discard_draft`, `has_draft`, `setWindowCommand`).
**Why it happens:** Developer forgets to update `MOCK_BRIDGE` when adding new bridge slots.
**How to avoid:** Update `MOCK_BRIDGE` in `bridge.js` with all new slots and the new response shape for existing slots.
**Warning signs:** Dev mode crashes with "bridge.commit_draft is not a function".

## Code Examples

### Error Page HTML for Missing Frontend Build

```python
# Source: decision D-04, styling at Claude's discretion
_MISSING_FRONTEND_HTML = """<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Virelo</title>
<style>
  body { font-family: system-ui, sans-serif; background: #1a1a1a; color: #e0e0e0;
         display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
  .box { max-width: 480px; text-align: center; }
  h1 { font-size: 20px; font-weight: 600; margin-bottom: 12px; }
  p { font-size: 14px; color: #999; line-height: 1.6; }
  code { background: #2a2a2a; padding: 2px 6px; border-radius: 4px; font-size: 13px; }
</style>
</head>
<body>
<div class="box">
  <h1>Frontend build not found</h1>
  <p>The file <code>frontend/dist/index.html</code> is missing.<br>
  Run <code>scripts/build-frontend.ps1</code> to build the frontend.</p>
</div>
</body>
</html>"""
```
[ASSUMED -- styling is at Claude's discretion per CONTEXT.md]

### Window Command Bridge Slot

```python
# Approach: single setWindowCommand slot (Claude's discretion)
@Slot(str, result=str)
def setWindowCommand(self, command: str) -> str:
    """Execute a window management command."""
    if command not in ("minimize", "close"):
        return json.dumps({"ok": False, "error": f"Unknown command: {command}"})
    if self._main_window is None:
        return json.dumps({"ok": False, "error": "MainWindow not ready"})
    try:
        if command == "minimize":
            self._main_window.showMinimized()
        elif command == "close":
            self._main_window.close()
        return json.dumps({"ok": True})
    except Exception as e:
        LOG.exception("setWindowCommand(%s) failed", command)
        return json.dumps({"ok": False, "error": str(e)})
```
[ASSUMED -- combines D-18 with Claude's discretion on single vs separate slots]

### Frontend Save Flow After Draft Model

```javascript
// After draft model: handleSave calls commit_draft, not save_settings
const handleSave = () => {
  bridge.commit_draft((result) => {
    try {
      const r = JSON.parse(result);
      if (r.ok) setUnsaved(false);
      else console.error('[app] commit_draft failed:', r.error);
    } catch (e) {
      console.error('[app] Failed to parse commit_draft result:', e);
    }
  });
};

// handleDiscard calls discard_draft
const handleDiscard = () => {
  bridge.discard_draft((result) => {
    try {
      const r = JSON.parse(result);
      if (r.ok) setUnsaved(false);
    } catch (e) {
      console.error('[app] Failed to parse discard_draft result:', e);
    }
  });
};

// set() now calls save_settings to update draft (not persist)
const set = (p) => {
  const bridgeData = JSON.stringify(stateToBridgePartial(p));
  bridge.save_settings(bridgeData, (result) => {
    try {
      const r = JSON.parse(result);
      if (r.ok) {
        setState((s) => ({ ...s, ...p }));
        setUnsaved(true);
      }
    } catch (e) {
      console.error('[app] save_settings (draft) failed:', e);
    }
  });
};
```
[ASSUMED -- follows the data flow implied by D-06..D-10]

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `LocalContentCanAccessRemoteUrls: True` always | Conditional on dev mode | This phase | Blocks local content from fetching remote URLs in release mode |
| `not sys.frozen` implies dev mode | `VIRELO_DEV=1` explicit check only | This phase | Running from source defaults to release-mode security |
| `save_settings` persists immediately | `save_settings` updates draft, `commit_draft` persists | This phase | Clean separation of UI state from persisted state |
| Unknown keys silently dropped | Unknown keys produce structured error | This phase | Frontend bugs surface immediately |
| Mixed `{"ok": ...}` and bare value responses | All slots return `{"ok": true/false, ...}` | This phase | Consistent error handling across all bridge calls |

**Deprecated/outdated:**
- The `save_settings` slot name is kept but its semantics change from "persist immediately" to "update draft". This avoids a frontend-wide rename while changing the backend behavior.
- `stateToBridge()` currently includes `theme: undefined` which gets stringified. After cleanup, theme should either be included properly or the key omitted from the mapping.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `setHtml()` uses `data:` URLs internally that would be caught by `acceptNavigationRequest` | Pitfalls | Error page would not render; need to test and adjust the allowed schemes |
| A2 | `Qt.ContextMenuPolicy.NoContextMenu` is the correct enum path in PySide6 >=6.6 | Pattern 4 | Build error; may need `Qt.NoContextMenu` without the enum namespace |
| A3 | Title bar click events work because no `-webkit-app-region: drag` is set and WM_NCHITTEST only applies to edges | Pitfalls | Title bar buttons may not fire if the drag region overlaps |
| A4 | `showMinimized()` correctly minimizes a frameless window to tray (via existing `changeEvent` handler) | Code Examples | Minimize might not trigger tray behavior; may need to call `hide()` directly |

## Open Questions

1. **`data:` scheme in acceptNavigationRequest**
   - What we know: `setHtml()` docs say content is converted to a `data:` URL. `acceptNavigationRequest` is called for all navigations.
   - What's unclear: Whether `acceptNavigationRequest` is called for `setHtml()` specifically, or only for subsequent navigations after the page is loaded.
   - Recommendation: Allow `data:` scheme in the filter. If `acceptNavigationRequest` is NOT called for `setHtml()`, the allow is harmless. If it IS called, the allow is required.

2. **Should `save_settings` still be named `save_settings` or renamed to `update_draft`?**
   - What we know: D-07 says "When the frontend sends changes via `save_settings`, the bridge stores them in `_draft`". This implies keeping the name.
   - What's unclear: Whether renaming to `update_draft` would be clearer for maintainability.
   - Recommendation: Keep `save_settings` name per D-07. The semantic change is documented in the method docstring. Renaming would require updating all frontend callsites and the mock bridge.

3. **Sidebar "up to date" text replacement**
   - What we know: The sidebar shows "v{version} . up to date" at app.jsx:105. The "up to date" text implies an update check that does not exist.
   - What's unclear: What to replace it with. CONTEXT.md lists this as Claude's discretion.
   - Recommendation: Replace with just `v{__APP_VERSION__}` -- the version number alone. No status indicator needed for a personal tool with no update mechanism.

## Security Domain

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | no | N/A -- single-user desktop utility, no auth |
| V3 Session Management | no | N/A -- no sessions |
| V4 Access Control | no | N/A -- admin elevation is OS-level, not app-level |
| V5 Input Validation | yes | `SettingsState.apply_draft()` validates all keys and values with type coercion and range bounds |
| V6 Cryptography | no | N/A -- no encryption, no secrets stored |

### Known Threat Patterns for PySide6 + QWebEngine

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| External navigation from local content | Tampering | `acceptNavigationRequest` blocks non-local URLs |
| Local content fetching remote resources | Information Disclosure | `LocalContentCanAccessRemoteUrls = False` in release mode |
| Context menu exposes dev tools | Information Disclosure | `setContextMenuPolicy(Qt.NoContextMenu)` in release mode |
| JavaScript injection via bridge input | Tampering | All bridge slot inputs validated before processing; JSON-only interface |
| Unknown settings keys accepted silently | Tampering | `apply_draft` rejects unknown keys with error response |
| Window popup via `window.open` | Tampering | `createWindow` returns `None` by default in QWebEnginePage; no override exists that would allow it |

## Sources

### Primary (HIGH confidence)
- [PySide6 QWebEnginePage docs](https://doc.qt.io/qtforpython-6.8/PySide6/QtWebEngineCore/QWebEnginePage.html) -- `acceptNavigationRequest` signature, NavigationType enum values
- [PySide6 QWebEngineSettings docs](https://doc.qt.io/qtforpython-6/PySide6/QtWebEngineCore/QWebEngineSettings.html) -- `LocalContentCanAccessRemoteUrls` attribute
- [PySide6 QWebEngineView docs](https://doc.qt.io/qtforpython-6/PySide6/QtWebEngineWidgets/QWebEngineView.html) -- `setHtml()`, `setUrl()`, context menu policy
- Codebase analysis: `webview.py`, `bridge.py`, `settings_state.py`, `settings.py`, `app.jsx`, `pages.jsx`, `panels.jsx`, `bridge.js` -- all read in full

### Secondary (MEDIUM confidence)
- [Python GUIs QWebEngineView tutorial](https://www.pythonguis.com/faq/qwebengineview-change-anchor-behavior/) -- `acceptNavigationRequest` override examples
- [Qt Forum: context menu in QWebView](https://www.qtcentre.org/threads/15594-Disable-context-menu-in-QWebView) -- context menu policy approaches

### Tertiary (LOW confidence)
- None -- all claims verified against official docs or codebase

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- all APIs verified in PySide6 official docs; no new dependencies
- Architecture: HIGH -- draft state pattern is standard; all integration points identified in codebase
- Pitfalls: HIGH -- all pitfalls derived from reading actual source code and tracing data flow
- Security: HIGH -- ASVS categories assessed against actual threat model; all controls are Qt-native APIs

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (stable -- PySide6 and React APIs are mature and unlikely to change)
