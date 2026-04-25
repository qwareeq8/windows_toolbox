# Phase 6: UI Action Placement - Context

**Gathered:** 2026-04-25 (auto mode)
**Status:** Ready for planning

<domain>
## Phase Boundary

Every visible control lives on the correct page, does what it claims, and nothing fake remains in the interface. Test Snap moves to the snap page, Reset Defaults moves to General with a confirmation dialog, the footer is cleaned up, key capture uses press-to-capture controls, and the command palette shows only real actions.

</domain>

<decisions>
## Implementation Decisions

### Test Snap relocation (UI-01, UI-02)
- **D-01:** Remove the Test Snap button from the Footer component. It currently sits as the first element in the footer — remove it entirely from Footer's render and props.
- **D-02:** Add Test Snap as a ghost Button in the Target Size card header. The Card component already supports a header bar — add the button right-aligned in the Target Size card's header div, next to the "Target size" title. Use `variant="ghost"` with the play icon, matching the existing style.
- **D-03:** Test Snap in the command palette stays — it's a valid action. It already calls `onTestSnap` which works correctly.
- **D-04:** Test Snap should use current draft values (not last-saved). Currently `handleTestSnap` calls `bridge.test_snap` which snaps using the draft overlay from `get_all()` — this is already correct since Phase 5's draft model returns draft-overlaid values. No change needed for UI-02.

### Footer cleanup (UI-03)
- **D-05:** Remove the `onTestSnap` prop from Footer. Remove the Test Snap Button render.
- **D-06:** Remove the `onReset` prop and Reset Defaults button from Footer. The footer should contain only: status message (left), flex spacer, dirty indicator + Discard (when unsaved), Save changes (always visible).
- **D-07:** The footer layout stays as a horizontal flex bar with `12px 24px` padding and a top border, matching the current design.

### Reset Defaults confirmation (UI-04)
- **D-08:** Move Reset Defaults to the GeneralPage under an "Advanced" card. The current GeneralPage already has an "Advanced" card with a "Reset all settings" row and a danger Button — wire the button's `onClick` to open a confirmation dialog.
- **D-09:** The confirmation dialog uses a custom overlay pattern similar to CommandPalette (position: absolute, backdrop, centered card). Not a native `confirm()` call — the project uses React-rendered UI throughout.
- **D-10:** Confirmation dialog content: "Reset all settings to defaults? This cannot be undone." with "Cancel" (secondary) and "Reset" (danger) buttons. On confirm, call `bridge.reset_defaults` and close the dialog.
- **D-11:** Remove `onReset` from VireloApp's Footer props. Keep `handleReset` in VireloApp but wire it through `GeneralPage` via `app.reset` or similar pattern. Alternatively, call `bridge.reset_defaults` directly in the GeneralPage confirmation handler.

### Key capture controls (UI-05)
- **D-12:** Replace the SHIFT/CTRL/ALT Segmented controls on the SnapPage with press-to-capture buttons. Each binding (snap key, restore key) shows a Kbd-styled button displaying the current key name.
- **D-13:** When the user clicks the button, it enters capture mode: the button highlights (accent border), text changes to "Press a key...", and the next keypress is captured via `bridge.capture_key(target, callback)`. The captured key shows as the new binding and triggers `dirty_changed` (already handled by Phase 5's draft model).
- **D-14:** Cancel capture: pressing Escape or clicking outside the capture button cancels without changing the binding. The button reverts to showing the current key.
- **D-15:** The capture control is a new component (e.g., `KeyCapture`) in pages.jsx or primitives.jsx. It takes `value` (current key name), `onCapture` callback, and uses local state for capture mode.

### Command palette truthfulness (UI-07)
- **D-16:** Theme toggle command currently uses `setTweaks({ theme: ... })` — update to use `app.set({ themeMode: ... })` to route through the bridge, consistent with Phase 5 changes. Include "System" as an option alongside Light/Dark.
- **D-17:** Accent commands currently use `setTweaks({ accent: ... })` — update to use `app.set({ accent: ... })` to route through the bridge.
- **D-18:** Remove the "Reset to defaults" command from the command palette — it now requires a confirmation dialog and belongs only on the General page (D-08). Triggering reset from the palette would bypass the confirmation.
- **D-19:** All remaining commands (navigation, test snap, save, enable/disable snap, game mode) are real and correctly wired — no changes needed.

### Shortcuts page accuracy (UI-06)
- **D-20:** Change the ShortcutsPage subtitle from "Click any to rebind." to "Global keyboard shortcuts registered by Virelo." The current hover interaction on shortcut rows is misleading — clicking doesn't actually rebind.
- **D-21:** Remove the hover effect and cursor:pointer from shortcut rows. They should be informational, not interactive. Key rebinding happens on the Snap page via the new capture controls (D-12).
- **D-22:** Ensure the Shortcuts page accurately shows the current bindings from `app.snapKey` and `app.restoreKey` — this already works. The Ctrl+K command palette shortcut is also accurate.

### Claude's Discretion
- Whether to extract the confirmation dialog as a reusable `ConfirmDialog` primitive or inline it in GeneralPage
- Whether the KeyCapture component lives in primitives.jsx or is defined locally in pages.jsx
- Animation/transition on capture mode state change
- Whether the "Reset to defaults" palette command is removed entirely or shows the confirmation dialog inline

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Vision, constraints, forbidden changes (no fake controls, no stale naming)
- `.planning/REQUIREMENTS.md` — UI-01..07 requirement definitions (Phase 6 scope)
- `CLAUDE.md` — Forbidden changes, known footguns

### Prior phase context
- `.planning/phases/05-bridge-and-settings-correctness/05-CONTEXT.md` — All Phase 5 decisions (dirty state, draft/commit, theme, key capture draft model)
- `.planning/phases/05-bridge-and-settings-correctness/05-RESEARCH.md` — Phase 5 research (pitfalls, patterns)
- `.planning/phases/02-security-and-bridge-hardening/02-CONTEXT.md` — Draft state architecture, fake control removal decisions

### Current implementation (files being modified)
- `frontend/src/app.jsx` — VireloApp: Footer props, handleReset, handleTestSnap
- `frontend/src/pages.jsx` — SnapPage snap/restore key controls, GeneralPage Reset button, ShortcutsPage subtitle
- `frontend/src/panels.jsx` — CommandPalette: theme/accent commands, reset command
- `frontend/src/primitives.jsx` — Card, Button, Kbd, Segmented components (reuse patterns)
- `frontend/src/bridge.js` — MOCK_BRIDGE capture_key mock (for dev mode testing)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `Card` component with header bar (`title`, `subtitle` props) — Test Snap button goes in this header
- `Button` component with `variant="ghost"` and `icon` prop — matches Test Snap button needs
- `Kbd` component — styling reference for key capture buttons
- `CommandPalette` overlay pattern (backdrop + centered card) — confirmation dialog can reuse this pattern
- `bridge.capture_key(target, callback)` slot — already exists and works, returns the captured key name

### Established Patterns
- All controls route through `app.set()` → `bridge.save_settings()` → `apply_draft()` (Phase 5)
- `dirty_changed` signal drives footer dirty indicator (Phase 5)
- Card headers use `surface2` background with `borderBottom` — consistent placement for Test Snap
- Ghost buttons used for secondary actions in the UI

### Integration Points
- `Footer` component in app.jsx needs prop removal (onTestSnap, onReset)
- `VireloApp` component needs `handleReset` wired through `app` object or removed from top level
- `SnapPage` key controls replace Segmented with KeyCapture component
- `CommandPalette` commands array needs theme/accent routing fix and reset removal
- `GeneralPage` Advanced card button needs onClick handler and confirmation dialog

</code_context>

<specifics>
## Specific Ideas

- The Test Snap relocation is primarily a Frontend-only change — the bridge `test_snap` slot works correctly and uses draft values already.
- Key capture (D-12..D-15) is the most complex new component in this phase. The `bridge.capture_key(target, cb)` API already exists — the frontend just needs a proper capture UX.
- The command palette theme toggle (D-16) should cycle through System/Light/Dark, not just toggle Light/Dark, since Phase 5 added System support.
- Footer cleanup (D-05..D-07) is straightforward removal — the footer becomes simpler.

</specifics>

<deferred>
## Deferred Ideas

None — analysis stayed within phase scope

</deferred>

---

*Phase: 06-ui-action-placement*
*Context gathered: 2026-04-25*
