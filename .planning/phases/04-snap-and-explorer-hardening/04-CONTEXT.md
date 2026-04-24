# Phase 4: Snap and Explorer Hardening - Context

**Gathered:** 2026-04-24
**Status:** Ready for planning

<domain>
## Phase Boundary

Harden the snap and Explorer features into isolated, tested services with correct scope and robust edge-case handling. Separate hotkey detection from window movement in the snap engine, exclude Virelo's own window from snapping, extract Explorer worker orchestration from MainWindow into a dedicated service, and add unit tests covering multi-monitor snap geometry scenarios.

</domain>

<decisions>
## Implementation Decisions

### Hotkey-snap separation
- **D-01:** Split `ShiftSnapRestore` into two concerns: a `HotkeyListener` class that detects multi-press keyboard patterns and emits a trigger signal, and the existing `ShiftSnapRestore` class which retains window movement logic (`_snap`, `_restore`, `perform`). The listener owns `keyboard.on_press_key`/`on_release_key` hooks, the press-time deque, and interval logic. ShiftSnapRestore receives the trigger signal and performs the window operation.
- **D-02:** Both classes stay in `virelo/services/snap.py` — no new module needed. HotkeyListener is an implementation detail, not a public API. SnapService facade remains the external interface for bridge.py.
- **D-03:** HotkeyListener emits a `triggered(bool)` signal (bool = restore modifier held). ShiftSnapRestore connects to this signal in its constructor or via MainWindow wiring. This matches the existing signal pattern.

### Virelo window exclusion
- **D-04:** When snap targets Virelo's own window (detected via Qt top-level widget check), skip entirely — return without moving or centering. SNAP-03 says "excluded from snapping," which means no action at all. Remove the current `center_on_screen()` call in `_snap()` for Qt widgets.
- **D-05:** The restore path should also skip Virelo's own window to stay consistent — if we never snap it, there's nothing to restore.
- **D-06:** Log a debug message when Virelo's window is skipped, for diagnostics.

### Explorer service extraction
- **D-07:** Create `ExplorerService` class in `virelo/services/explorer_service.py` (or add to existing `explorer_columns.py` if it fits). This service owns the explorer worker lifecycle: start, stop, and the `_autosize_explorer_columns_quick`/`_autosize_explorer_columns_full` wrapper functions currently in `window.py`.
- **D-08:** ExplorerService follows the same facade pattern as SnapService — MainWindow creates it, passes it any needed references, and delegates start/stop to it. MainWindow's `_update_explorer_autosize_thread`, `_stop_explorer_worker`, and `_on_explorer_finished` move into ExplorerService.
- **D-09:** ExplorerService starts the worker only when `ex_auto_size` setting is enabled (EXPL-03). It exposes `start()`, `stop()`, and `is_running()` methods. The bridge or MainWindow calls these on setting changes.
- **D-10:** COM threading constraint (D-07 from Phase 3) still applies: ExplorerAutosizeWorker's COM init and COM operations must stay co-located in `workers/explorer.py`. The new ExplorerService manages the QThread lifecycle but does not touch COM internals.

### Explorer page scope
- **D-11:** EXPL-01 is already satisfied — Phase 2 removed fake controls. The Explorer page shows only the auto-size columns toggle. No changes needed to the frontend for this requirement.

### Test coverage
- **D-12:** Add `test_calculate_snap_position_negative_coords` — snap geometry on a monitor with negative x/y origin (left-of-primary layout). This fills the SNAP-05 gap for negative coordinates in the snap calculation (existing `test_rect_matches_negative_coords` only tests fullscreen detection).
- **D-13:** Add `test_calculate_snap_position_vertical_layout` — snap on a monitor below the primary (y offset). Covers vertical multi-monitor layouts.
- **D-14:** Add tests for the maximized-restore round-trip: verify that `_restore()` logic correctly detects `was_maximized` flag and issues `SW_MAXIMIZE` (unit-testable by mocking win32gui calls). This validates SNAP-04.
- **D-15:** Existing unit tests in `tests/unit/test_snap_geometry.py` already cover single monitor and two-monitor horizontal. Extend this file with the new scenarios rather than creating a new test file.

### Claude's Discretion
- Internal structure of HotkeyListener (whether it subclasses QObject or uses plain Python with a callback)
- Whether ExplorerService lives in its own file or is added to explorer_columns.py
- Exact mock strategy for win32gui in restore tests (conftest fixtures vs inline mocks)
- Whether to add a `blocked` signal to SnapService for game-mode skip notifications

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Project context
- `.planning/PROJECT.md` — Vision, constraints, out-of-scope features
- `.planning/REQUIREMENTS.md` — SNAP-01..05, EXPL-01..03 requirement definitions
- `CLAUDE.md` — Forbidden changes, known footguns

### Prior phase context
- `.planning/phases/03-structure-and-quality/03-CONTEXT.md` — Package layout (D-02), COM threading constraint (D-07), explorer_columns.py size exception (D-04), test tiering (D-08..D-10)
- `.planning/phases/02-security-and-bridge-hardening/02-CONTEXT.md` — Explorer page cleanup (D-14), bridge response standardization

### Architecture
- `.planning/codebase/ARCHITECTURE.md` — Current architecture layers and data flow
- `.planning/codebase/STRUCTURE.md` — Current file layout
- `.planning/codebase/CONVENTIONS.md` — Naming, imports, error handling patterns

### Current implementation (files being modified)
- `virelo/services/snap.py` — SnapService facade + ShiftSnapRestore engine (split target for D-01..D-03)
- `virelo/app/window.py` — MainWindow with explorer orchestration (extraction target for D-07..D-10)
- `virelo/workers/explorer.py` — ExplorerAutosizeWorker (COM threading, not modified but referenced)
- `virelo/services/explorer_columns.py` — COM column manager (not modified)
- `virelo/platform/win32_helpers.py` — DPI, monitor rect, fullscreen detection helpers
- `tests/unit/test_snap_geometry.py` — Existing snap geometry tests (extended with D-12..D-15)
- `tests/conftest.py` — Native module stubs for unit testing without PySide6/Win32
- `frontend/src/pages.jsx` — Explorer page (EXPL-01 already satisfied, no changes needed)

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `SnapService` facade in `snap.py` — Already provides narrow API for bridge. HotkeyListener will be an internal detail behind this facade.
- `calculate_snap_position()` pure function in `snap.py` — Already extracted and testable. No changes needed.
- `_autosize_explorer_columns_quick/full` wrappers in `window.py` — Move to ExplorerService.
- Test fixtures in `tests/conftest.py` — Native module stubs (win32gui, win32con, keyboard, etc.) enable unit testing without Win32 APIs.

### Established Patterns
- SnapService facade pattern — ExplorerService should follow the same pattern (thin wrapper, lifecycle management, structured error returns)
- Signal-slot for cross-thread communication — HotkeyListener.triggered signal connects to ShiftSnapRestore.perform slot
- Bottom-up extraction from Phase 3 — move pure logic first, wire last
- Two-tier testing (unit/integration) from Phase 3 D-08

### Integration Points
- `MainWindow.__init__` creates ShiftSnapRestore and wires signals — will also create ExplorerService
- `bridge.py` delegates to SnapService — no changes needed if SnapService API doesn't change
- `bridge._apply_side_effects()` starts/stops explorer worker — will delegate to ExplorerService instead

</code_context>

<specifics>
## Specific Ideas

- The hotkey separation is primarily an architecture improvement for testability. The user-visible behavior should not change — same multi-press pattern, same snap/restore behavior.
- Virelo window exclusion is a behavioral fix: currently snapping while Virelo is focused centers Virelo, which is unexpected. After D-04, snapping while Virelo is focused simply does nothing.
- ExplorerService extraction mirrors the SnapService pattern established in Phase 3, maintaining consistency across the codebase.
- EXPL-01 requires no work — verified that Phase 2 already cleaned up the Explorer page to show only auto-size columns.

</specifics>

<deferred>
## Deferred Ideas

None — discussion stayed within phase scope

</deferred>

---

*Phase: 04-snap-and-explorer-hardening*
*Context gathered: 2026-04-24*
