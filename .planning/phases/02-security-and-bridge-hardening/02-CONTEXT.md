# Phase 2: Security and Bridge Hardening - Context

**Gathered:** 2026-04-24
**Status:** Ready for planning

<domain>
## Phase Boundary

Lock down the WebEngine host for an admin-elevated process, restructure the bridge with a Python-side draft/commit state model, and ensure every visible frontend control connects to a real Python backend method. Remove all fake controls that have no backend implementation.

</domain>

<decisions>
## Implementation Decisions

### WebEngine lockdown
- **D-01:** Override `acceptNavigationRequest` in `VireloWebPage` to block all non-local navigation. Allow `file://` and `http://localhost` (dev server) URLs only. Log blocked navigation attempts.
- **D-02:** `LocalContentCanAccessRemoteUrls` set to `False` in release mode, `True` only when `_is_dev_mode()` returns True.
- **D-03:** Dev mode detection must require `VIRELO_DEV=1` explicitly — remove the `not sys.frozen` fallback. Running from source without the env var should behave like release mode (with local files, not Vite dev server).
- **D-04:** When `frontend/dist/index.html` is missing in release mode, display a styled inline HTML error page (not a blank white page) explaining the build is missing and how to run `scripts/build-frontend.ps1`.
- **D-05:** Default WebEngine context menu disabled in release mode. Enabled in dev mode for debugging (Inspect Element access).

### Draft state architecture
- **D-06:** Add a `_draft` dict to `SettingsState` that holds unsaved changes separately from the persisted `Settings` object. On construction, `_draft` is `None` (no pending changes).
- **D-07:** When the frontend sends changes via `save_settings`, the bridge stores them in `_draft` instead of immediately writing to QSettings. The frontend reflects draft values.
- **D-08:** A new `commit_draft` bridge slot persists `_draft` to QSettings via `Settings.save()`, applies side effects (startup shortcut, snap manager, Explorer worker), then clears `_draft`.
- **D-09:** A new `discard_draft` bridge slot clears `_draft` and pushes the persisted settings back to the frontend via `settings_changed` signal.
- **D-10:** `get_settings` returns the merged view: persisted settings overlaid with any draft values. A separate `has_draft` slot or signal indicates whether unsaved changes exist.

### Fake control removal
- **D-11:** Remove these controls entirely from the frontend (no stubs, no placeholders):
  - Explorer page: "Remember column widths" toggle, "Hidden files" card (show hidden files + show file extensions toggles)
  - General page: "Start minimized to tray" toggle, "Automatic updates" toggle, "Anonymous telemetry" toggle
  - About page: "Check for updates" button, "Up to date" badge, "Documentation Open" button, "Report an issue Open" button
- **D-12:** Remove corresponding fake state keys from `bridgeToState` and `stateToBridge`: `rememberCols`, `showHidden`, `showExts`, `startTray`, `autoUpdate`, `telemetry`.
- **D-13:** Remove hardcoded changelog from About page — version number from `__APP_VERSION__` is sufficient. The About page should show: app icon, name, version, and license info only.
- **D-14:** Explorer page retains only the auto-size columns toggle after cleanup.

### Bridge response standardization
- **D-15:** All bridge slots return structured JSON: `{"ok": true, "data": ...}` on success, `{"ok": false, "error": "..."}` on failure. This applies to `get_settings`, `reset_defaults`, `test_snap`, `capture_key`, `apply_theme`, `get_theme_mode`, `toggle_run_at_startup`, `get_launch_at_login`, `get_snap_enabled`.
- **D-16:** Unknown keys in `apply_partial` are rejected with a structured error response `{"ok": false, "error": "Unknown keys: [list]"}` instead of being silently dropped. This surfaces frontend bugs early.
- **D-17:** The bridge validates all inputs before acting. Current validation in `capture_key` and `apply_theme` is correct — extend this pattern to all slots.

### Title bar and command palette wiring
- **D-18:** Title bar minimize (`—`) and close (`×`) buttons call bridge slots `setWindowCommand("minimize")` and `setWindowCommand("close")`. The maximize button (`▢`) is removed — frameless window with no maximize behavior.
- **D-19:** Command palette "Test snap" action calls `bridge.test_snap()` (currently a no-op `() => {}`).
- **D-20:** Command palette "Reset to defaults" action calls the same `handleReset` used by the footer button (currently a no-op `() => {}`).
- **D-21:** Command palette "Save changes" action calls `handleSave` instead of the current `app.set({ _saved: true })` which does nothing useful.

### Shortcuts page cleanup
- **D-22:** Remove fake shortcut entries that don't exist in the backend: "Cycle preset sizes", "Move to next display", "Move to previous display", "Toggle Virelo pause", "Open settings". Keep only: "Trigger snap" and "Restore last snap" (which are real) and "Command palette" (Ctrl+K, which is real frontend functionality).

### Claude's Discretion
- Error page HTML styling and exact copy
- Whether `get_settings` returns draft-merged or persisted-only (recommended: draft-merged)
- Internal implementation of `_draft` (dict overlay vs full copy)
- Whether to add a `setWindowCommand` bridge slot or separate `minimizeWindow`/`closeWindow` slots
- Sidebar "up to date" status text replacement after fake update status removal

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Project vision, constraints, key decisions, out-of-scope features
- `.planning/REQUIREMENTS.md` — SEC-01..05, BRDG-01..06, UI-01..06 requirement definitions
- `CLAUDE.md` — Forbidden changes (no fake controls, no stale naming), known footguns

### Architecture
- `.planning/codebase/ARCHITECTURE.md` — Full architecture description, data flow, bridge pattern
- `.planning/codebase/CONCERNS.md` — Bridge/MainWindow coupling, security considerations

### Phase 1 context
- `.planning/phases/01-hygiene-and-build-pipeline/01-CONTEXT.md` — Prior decisions (version consolidation, build pipeline)

### Current implementation files
- `webview.py` — WebEngine host (security settings, dev mode detection, navigation)
- `bridge.py` — VireloBridge (all bridge slots, signals, side effects)
- `settings.py` — Settings persistence (QSettings read/write)
- `settings_state.py` — SettingsState (JSON facade, validation, apply_partial)
- `frontend/src/app.jsx` — App shell, bridgeToState/stateToBridge mappings, save/discard/reset handlers
- `frontend/src/pages.jsx` — All settings pages (fake controls located here)
- `frontend/src/panels.jsx` — Command palette (no-op actions located here)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `VireloBridge` class in `bridge.py` — Already has the Slot/Signal pattern. New slots (`commit_draft`, `discard_draft`, `setWindowCommand`) follow the same pattern.
- `SettingsState` in `settings_state.py` — Already has `apply_partial` with validation. Draft model extends this class.
- `VireloWebPage` in `webview.py` — Already subclasses `QWebEnginePage` for console routing. Override `acceptNavigationRequest` here.
- Footer component in `app.jsx` — Already has Save/Discard/Reset buttons. Wire to new draft-aware bridge slots.
- `_apply_side_effects` in `bridge.py` — Side effect dispatch already exists. Move it to fire on `commit_draft` instead of `save_settings`.

### Established Patterns
- Bridge slots return JSON strings, signals push JSON strings — all new slots must follow this
- `_safe_bool` / `_safe_int` in `settings.py` — defensive type coercion pattern
- `bridgeToState` / `stateToBridge` in `app.jsx` — key mapping between Python snake_case and React camelCase
- `_is_dev_mode()` in `webview.py` — environment detection function (to be tightened)

### Integration Points
- `MainWindow.__init__` in `main.py` — Creates Settings, SettingsState, VireloBridge, VireloWebView in sequence
- `bridge.set_main_window(mw)` — Post-construction wiring for window operations (minimize/close will use this)
- `app.jsx` `handleSave`/`handleDiscard`/`handleReset` — Frontend save flow (to be rewired to draft model)
- `panels.jsx` command list — Static array of commands (no-ops to be replaced with real handlers)

</code_context>

<specifics>
## Specific Ideas

- The draft model should feel invisible to the user — they change settings, see changes reflected immediately in the UI, and Save/Discard commits or reverts. Same UX as today but with proper state separation underneath.
- Error page for missing frontend should be minimal and developer-friendly, not user-facing polish. This is a build error, not a user error.
- The Shortcuts page after cleanup will be very short (3 items). That's fine — it accurately represents what exists.

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 02-security-and-bridge-hardening*
*Context gathered: 2026-04-24*
