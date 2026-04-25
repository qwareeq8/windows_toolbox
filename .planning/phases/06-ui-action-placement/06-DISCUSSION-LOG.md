# Phase 6: UI Action Placement - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-25
**Phase:** 06-ui-action-placement
**Areas discussed:** Test Snap placement, Footer restructuring, Reset Defaults confirmation, Key capture controls, Command palette cleanup, Shortcuts page accuracy
**Mode:** auto (all areas auto-selected, recommended options auto-chosen)

---

## Test Snap Placement

| Option | Description | Selected |
|--------|-------------|----------|
| Ghost button in card header | Right-aligned in Target Size card header bar | ✓ |
| Standalone card | Separate card below Target Size with Test Snap action | |
| Floating action button | Fixed position button in bottom corner | |

**User's choice:** Ghost button in card header (auto-selected, recommended default)
**Notes:** Matches existing Card header pattern. Card component already has header bar support.

---

## Footer Restructuring

| Option | Description | Selected |
|--------|-------------|----------|
| Minimal footer | Status, dirty indicator, Discard, Save only | ✓ |
| Keep Test Snap in footer | Move only Reset Defaults out | |

**User's choice:** Minimal footer (auto-selected, per UI-03 requirement)
**Notes:** Directly implements UI-03 requirement.

---

## Reset Defaults Confirmation

| Option | Description | Selected |
|--------|-------------|----------|
| Custom overlay dialog | React-rendered overlay matching CommandPalette style | ✓ |
| Native confirm() | Browser native confirmation dialog | |
| Two-step button | Click once to reveal danger button, click again to confirm | |

**User's choice:** Custom overlay dialog (auto-selected, recommended default)
**Notes:** Consistent with project's React-rendered UI approach. No native browser dialogs.

---

## Key Capture Controls

| Option | Description | Selected |
|--------|-------------|----------|
| Kbd-style capture button | Click to enter capture mode, shows "Press a key..." | ✓ |
| Inline text input | Text field that captures keypress | |
| Dropdown with common keys | Select from list of modifier keys | |

**User's choice:** Kbd-style capture button (auto-selected, recommended default)
**Notes:** Most intuitive for keyboard shortcut rebinding. Uses existing Kbd component as styling reference.

---

## Command Palette Cleanup

| Option | Description | Selected |
|--------|-------------|----------|
| Route through app.set, remove Reset | Fix theme/accent routing, remove Reset (needs confirmation) | ✓ |
| Keep all commands, add confirmation inline | Show confirmation in palette for Reset | |

**User's choice:** Route through app.set, remove Reset (auto-selected, recommended default)
**Notes:** Reset without confirmation is dangerous. Palette bypass would violate UI-04.

---

## Shortcuts Page Accuracy

| Option | Description | Selected |
|--------|-------------|----------|
| Informational only | Remove "Click any to rebind" hint, no hover interaction | ✓ |
| Add rebind from shortcuts page | Make clicking actually rebind keys | |

**User's choice:** Informational only (auto-selected, recommended default)
**Notes:** Rebinding happens on Snap page via key capture controls. Shortcuts page is reference only.

---

## Claude's Discretion

- Whether to extract ConfirmDialog as reusable primitive or inline in GeneralPage
- Whether KeyCapture component lives in primitives.jsx or pages.jsx
- Animation/transition details for capture mode
- Theme toggle cycle order in command palette (System/Light/Dark)

## Deferred Ideas

None — discussion stayed within phase scope
