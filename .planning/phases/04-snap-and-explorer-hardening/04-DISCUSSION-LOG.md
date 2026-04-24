# Phase 4: Snap and Explorer Hardening - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-04-24
**Phase:** 04-snap-and-explorer-hardening
**Areas discussed:** Hotkey-snap separation, Virelo window exclusion, Explorer service extraction, Test coverage scope
**Mode:** --auto (all decisions auto-selected)

---

## Hotkey-Snap Separation

| Option | Description | Selected |
|--------|-------------|----------|
| Split into HotkeyListener + ShiftSnapRestore | Separate keyboard detection into its own class, keep movement logic in ShiftSnapRestore. Both in snap.py. | ✓ |
| Extract hotkey into standalone module | Move hotkey detection to a new file (e.g., hotkey.py). More files, more import paths. | |
| Keep combined but add test seams | Add methods/parameters to make the combined class testable without splitting. | |

**User's choice:** [auto] Split into HotkeyListener + ShiftSnapRestore (recommended — minimal class proliferation, clean signal boundary)
**Notes:** Matches SNAP-01/02 requirements. Signal-based decoupling allows testing hotkey logic without window operations.

---

## Virelo Window Exclusion

| Option | Description | Selected |
|--------|-------------|----------|
| Skip entirely (no-op) | When snap targets Virelo's own window, return immediately without any action. | ✓ |
| Center Virelo's window (current behavior) | Keep the existing center_on_screen() call for Virelo's window. | |
| Make configurable | Add a setting for whether Virelo should be snapped/centered/skipped. | |

**User's choice:** [auto] Skip entirely (recommended — matches SNAP-03 "excluded" language)
**Notes:** Current behavior centers Virelo's window on snap, which is unexpected. "Excluded" means no action.

---

## Explorer Service Extraction

| Option | Description | Selected |
|--------|-------------|----------|
| ExplorerService facade (like SnapService) | Create a service class that manages worker lifecycle, consistent with existing pattern. | ✓ |
| Move orchestration into ExplorerAutosizeWorker | Worker manages its own thread lifecycle. Breaks existing QThread pattern. | |
| Keep in MainWindow, just add start/stop methods | Minimal change, but MainWindow stays bloated. | |

**User's choice:** [auto] ExplorerService facade (recommended — consistent with SnapService pattern)
**Notes:** EXPL-02 requires moving orchestration out of MainWindow. Facade pattern is already established.

---

## Test Coverage Scope

| Option | Description | Selected |
|--------|-------------|----------|
| Fill gaps + maximized-restore test | Add negative-coords and vertical-layout snap tests, plus restore round-trip test. | ✓ |
| Minimum required by SNAP-05 only | Only add negative-coords test for calculate_snap_position. | |

**User's choice:** [auto] Fill gaps + maximized-restore test (recommended — thorough coverage for SNAP-04 and SNAP-05)
**Notes:** Existing tests cover single monitor, offset monitor, and negative coords for fullscreen detection. Need negative coords for snap calculation and restore behavior validation.

---

## Claude's Discretion

- HotkeyListener internal structure (QObject subclass vs plain Python)
- ExplorerService file placement
- Mock strategy for win32gui in restore tests
- Whether to add blocked signal to SnapService

## Deferred Ideas

None — all discussion stayed within phase scope.
