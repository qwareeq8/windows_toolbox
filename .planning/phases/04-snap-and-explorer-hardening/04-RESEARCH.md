# Phase 4: Snap and Explorer Hardening - Research

**Researched:** 2026-04-24
**Domain:** Python refactoring / service extraction / Win32 window management / unit testing
**Confidence:** HIGH

## Summary

Phase 4 hardens two existing subsystems -- snap/restore and Explorer column auto-sizing -- by separating concerns, extracting services, fixing behavioral edge cases, and adding test coverage. All decisions are locked in CONTEXT.md (D-01 through D-15) and well-constrained. The work is entirely within existing Python code with no new dependencies, no frontend changes, and no architectural shifts.

The primary risk is signal/slot disconnection during the ShiftSnapRestore split (D-01). The snap engine currently entangles hotkey detection with window movement in a single class. Separating these into HotkeyListener and ShiftSnapRestore requires careful wiring of the `triggered(bool)` signal. The secondary risk is the ExplorerService extraction (D-07/D-08), which moves thread lifecycle management out of MainWindow while respecting the COM threading constraint (Phase 3 D-07).

**Primary recommendation:** Execute bottom-up -- pure logic changes first (Virelo window exclusion, test additions), then internal class splits (HotkeyListener extraction), then service extraction (ExplorerService), then wiring updates (MainWindow, bridge) last.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Split ShiftSnapRestore into HotkeyListener (keyboard pattern detection, emits trigger signal) and ShiftSnapRestore (window movement logic). Listener owns keyboard hooks, press deque, interval logic.
- **D-02:** Both classes stay in `virelo/services/snap.py`. HotkeyListener is internal. SnapService facade remains the external API.
- **D-03:** HotkeyListener emits `triggered(bool)` signal (bool = restore modifier held). ShiftSnapRestore connects to this signal.
- **D-04:** When snap targets Virelo's own window, skip entirely -- no action, no center_on_screen(). Remove current centering behavior.
- **D-05:** Restore path also skips Virelo's own window for consistency.
- **D-06:** Log debug message when Virelo's window is skipped.
- **D-07:** Create ExplorerService in `virelo/services/explorer_service.py` (or add to existing explorer_columns.py). Owns explorer worker lifecycle.
- **D-08:** ExplorerService follows SnapService facade pattern. MainWindow delegates start/stop to it.
- **D-09:** ExplorerService starts worker only when `ex_auto_size` setting is enabled. Exposes start(), stop(), is_running().
- **D-10:** COM threading constraint preserved -- ExplorerAutosizeWorker COM init and operations stay in workers/explorer.py. ExplorerService manages QThread lifecycle only.
- **D-11:** EXPL-01 already satisfied by Phase 2. No frontend changes needed.
- **D-12:** Add test for snap geometry on monitor with negative x/y origin (left-of-primary layout).
- **D-13:** Add test for snap on monitor below the primary (y offset, vertical layout).
- **D-14:** Add tests for maximized-restore round-trip (verify was_maximized flag and SW_MAXIMIZE issuance).
- **D-15:** Extend existing tests/unit/test_snap_geometry.py rather than creating new test files.

### Claude's Discretion
- Internal structure of HotkeyListener (QObject subclass vs plain Python with callback)
- Whether ExplorerService lives in its own file or is added to explorer_columns.py
- Exact mock strategy for win32gui in restore tests (conftest fixtures vs inline mocks)
- Whether to add a `blocked` signal to SnapService for game-mode skip notifications

### Deferred Ideas (OUT OF SCOPE)
None -- discussion stayed within phase scope
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| SNAP-01 | Snap logic extracted from main.py into a dedicated service module | Already done in Phase 3 (snap.py exists in virelo/services/). D-01 further separates concerns within snap.py. |
| SNAP-02 | Hotkey detection separated from window movement logic | D-01/D-02/D-03: HotkeyListener class extracted from ShiftSnapRestore within snap.py |
| SNAP-03 | Virelo's own window excluded from snapping | D-04/D-05/D-06: Skip entirely in both _snap() and _restore() paths, log debug message |
| SNAP-04 | Restore correctly handles previously-maximized windows | D-14: Existing code already has was_maximized + SW_MAXIMIZE logic; test validates correctness |
| SNAP-05 | Geometry calculations have unit tests covering multi-monitor scenarios | D-12/D-13/D-15: Three new test scenarios added to existing test_snap_geometry.py |
| EXPL-01 | Explorer page shows only implemented features (auto-size columns) | D-11: Already satisfied by Phase 2 cleanup -- no changes needed |
| EXPL-02 | Explorer worker orchestration moved out of MainWindow into a service | D-07/D-08: ExplorerService extracted from MainWindow, follows SnapService pattern |
| EXPL-03 | Explorer worker starts only when setting is enabled and stops cleanly on quit | D-09/D-10: ExplorerService.start()/stop()/is_running() with setting gate |
</phase_requirements>

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Hotkey detection (multi-press pattern) | Backend (Python workers) | -- | keyboard library hooks run at OS level, must be Python-side |
| Window snap/restore | Backend (Python services) | -- | Win32 API calls (MoveWindow, ShowWindow) are Python-only |
| Virelo window exclusion | Backend (Python services) | -- | Qt widget detection is Python-side; no frontend awareness needed |
| Explorer worker lifecycle | Backend (Python services) | -- | QThread management is Python; COM stays in worker thread |
| Explorer page UI | Frontend (React) | -- | Already correct per EXPL-01 (Phase 2). No changes. |
| Snap geometry math | Backend (Python services) | -- | Pure function, no platform dependency. Primary test target. |

## Architecture Patterns

### System Architecture Diagram

```
Keyboard Event (OS)
        |
        v
  HotkeyListener (NEW)
  [keyboard hooks, press timing deque, interval logic]
        |
        | triggered(bool) signal
        v
  ShiftSnapRestore (MODIFIED)
  [_snap(), _restore(), perform()]
        |
        +--> Virelo window? --> SKIP + LOG (D-04/D-05/D-06)
        |
        +--> External window --> Win32 MoveWindow / ShowWindow
                                  |
                                  +--> get_monitor_rect()
                                  +--> calculate_snap_position()
                                  +--> _is_window_fullscreen()


Settings Change ("ex_auto_size")
        |
        v
  ExplorerService (NEW)
  [start(), stop(), is_running()]
        |
        | QThread lifecycle
        v
  ExplorerAutosizeWorker (UNCHANGED)
  [COM STA, Shell.Application, iter_tabs()]
        |
        v
  ExplorerAutosizeEngine (UNCHANGED)
  [step(), debounce, circuit breaker]
```

### Recommended Project Structure Changes
```
virelo/services/
  snap.py             # MODIFIED: HotkeyListener + ShiftSnapRestore + SnapService
  explorer_service.py # NEW: ExplorerService (worker lifecycle facade)
  explorer_columns.py # UNCHANGED

virelo/app/
  window.py           # MODIFIED: delegates explorer lifecycle to ExplorerService

virelo/workers/
  explorer.py         # UNCHANGED (COM constraint)

tests/unit/
  test_snap_geometry.py  # EXTENDED: negative coords, vertical layout, maximized restore
```

### Pattern 1: HotkeyListener Extraction (D-01/D-02/D-03)

**What:** Extract keyboard hook management, press timing, and trigger logic from ShiftSnapRestore into a new HotkeyListener class within snap.py. [VERIFIED: virelo/services/snap.py lines 93-219]

**When to use:** When the class doing keyboard detection also does window movement -- separation of concerns.

**Current code to extract (from ShiftSnapRestore):**
- `__init__`: keyboard.on_press_key / on_release_key hooks, press_times deque, _held flag
- `_on_press()`: press timing logic, trigger threshold, restore key check
- `_on_release()`: _held reset
- `cleanup()`: keyboard.unhook calls
- `update_binding()`: rebind keyboard hooks
- `update_press_limit()`: resize deque
- `current_key`, `restore_key` attributes

**What stays in ShiftSnapRestore:**
- `perform(restore)`: entry point for snap/restore
- `_snap(hwnd)`: window movement logic
- `_restore(hwnd)`: window restoration logic
- `_fetch_open_windows()`: initial window scan
- `_prune_closed_windows()`: stale window cleanup
- `_orig_sizes`: original position storage
- `triggered` signal (moved to HotkeyListener, ShiftSnapRestore receives it)
- `blocked` signal (stays, emitted from _snap during game mode skip)

**Design recommendation (Claude's discretion):** Make HotkeyListener a QObject subclass. This is simpler than a plain Python class with callbacks because:
1. The `triggered(bool)` signal is already defined as a Qt Signal [VERIFIED: snap.py line 94]
2. The existing wiring pattern uses `shift_mgr.triggered.connect(shift_mgr.perform)` [VERIFIED: window.py line 242]
3. QObject enables thread-safe signal emission from keyboard hook callbacks

```python
# Recommended structure within snap.py
class HotkeyListener(QtCore.QObject):
    """Detects multi-press keyboard patterns and emits trigger signal."""
    triggered = QtCore.Signal(bool)  # bool = restore modifier held

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self._press_times = deque(maxlen=normalize_snap_presses(settings.snap_presses))
        self._press_lock = threading.Lock()
        self._held = False
        self.current_key = str(settings.snap_key)
        self.restore_key = str(getattr(settings, "restore_key", "ctrl"))
        self._press_hook = keyboard.on_press_key(self.current_key, self._on_press)
        self._release_hook = keyboard.on_release_key(self.current_key, self._on_release)

    def cleanup(self):
        # unhook press/release hooks

    def update_binding(self, new_key):
        # unhook old, hook new

    def update_restore_key(self, new_key):
        self.restore_key = new_key

    def update_press_limit(self, new_limit):
        # resize deque

    def _on_press(self, event):
        # press timing + threshold + trigger signal emission

    def _on_release(self, event):
        self._held = False


class ShiftSnapRestore(QtCore.QObject):
    """Performs window snap and restore operations."""
    blocked = QtCore.Signal(str)

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self._orig_sizes = {}
        self._fetch_open_windows()

    @QtCore.Slot(bool)
    def perform(self, restore):
        # same as current, minus keyboard logic

    def _snap(self, hwnd):
        # window movement (with Virelo exclusion per D-04)

    def _restore(self, hwnd):
        # window restoration (with Virelo exclusion per D-05)
```

### Pattern 2: ExplorerService Extraction (D-07/D-08/D-09)

**What:** Move explorer worker thread lifecycle from MainWindow into a dedicated ExplorerService class. [VERIFIED: window.py lines 424-496]

**Code to move from MainWindow:**
- `_autosize_explorer_columns_quick()` (lines 40-52) -- module-level function
- `_autosize_explorer_columns_full()` (lines 55-68) -- module-level function
- `_update_explorer_autosize_thread()` (lines 424-469) -- becomes `start()`
- `_stop_explorer_worker()` (lines 474-492) -- becomes `stop()`
- `_on_explorer_finished()` (lines 494-496) -- internal callback
- `_explorer_thread` and `_explorer_worker` attributes

**What stays in MainWindow:**
- `_update_explorer_enabled_state()` -- thin delegate to ExplorerService.start/stop

```python
# virelo/services/explorer_service.py
class ExplorerService:
    """Lifecycle manager for the Explorer column auto-size worker."""

    def __init__(self, settings, parent_qobject=None):
        self._settings = settings
        self._parent = parent_qobject
        self._thread = None
        self._worker = None

    def start(self):
        """Start the explorer worker if ex_auto_size is enabled."""
        if not bool(self._settings.ex_auto_size):
            self.stop()
            return
        if self._thread and self._thread.isRunning():
            return
        # create thread, worker, connect signals, start

    def stop(self):
        """Stop the explorer worker cleanly."""
        # worker.stop(), thread.quit(), thread.wait(3000)

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.isRunning()
```

### Pattern 3: Virelo Window Exclusion (D-04/D-05/D-06)

**What:** Replace the current `center_on_screen()` behavior for Virelo's own window with a no-op skip. [VERIFIED: snap.py lines 238-248 for _snap, lines 316-329 for _restore]

**Current behavior in `_snap()`:**
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        center_fn = getattr(widget, "center_on_screen", None)
        if center_fn is not None:
            center_fn()  # <-- D-04 says REMOVE this
            widget.raise_()
            widget.activateWindow()
        return
```

**New behavior:**
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        LOG.debug("Snap: skipping Virelo's own window hwnd=%s", hwnd)
        return  # Skip entirely per SNAP-03
```

**Current behavior in `_restore()`:**
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        orig = self._orig_sizes.pop(hwnd, None)
        if not orig:
            return
        # ... restores widget geometry
        return
```

**New behavior:**
```python
for widget in QtWidgets.QApplication.topLevelWidgets():
    if int(widget.winId()) == hwnd:
        LOG.debug("Restore: skipping Virelo's own window hwnd=%s", hwnd)
        return  # Skip entirely per D-05
```

### Anti-Patterns to Avoid

- **Moving COM operations out of workers/explorer.py:** The COM STA apartment must be initialized and used within the same thread. ExplorerService manages thread lifecycle but NEVER touches COM objects directly. [VERIFIED: Phase 3 D-07]
- **Breaking SnapService facade API:** bridge.py depends on SnapService's test_snap(), update_binding(), update_restore_key(), update_press_limit() methods. These must continue working unchanged after the internal split. [VERIFIED: bridge.py line 273-280, snap.py lines 50-86]
- **Making HotkeyListener a public API:** D-02 explicitly says HotkeyListener is an implementation detail. SnapService should be the only external interface. Consumers should not import or depend on HotkeyListener directly.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Keyboard hook management | Custom keyboard polling | `keyboard` library hooks | Already used; handles OS-level key events reliably |
| Thread-safe signal emission | Manual callback threading | Qt Signal/Slot + QObject | Qt marshals cross-thread signals to the receiver's thread |
| COM lifecycle management | Custom COM init/teardown | pythoncom CoInitializeEx + worker pattern | COM STA threading is fragile; existing pattern is proven |
| Window geometry calculations | New geometry library | `calculate_snap_position()` pure function | Already extracted, tested, correct |

## Common Pitfalls

### Pitfall 1: Signal Disconnection During Split
**What goes wrong:** After extracting HotkeyListener from ShiftSnapRestore, the `triggered` signal is no longer on the same object that performs the action. If wiring in MainWindow is not updated, snap stops working.
**Why it happens:** The current wiring is `shift_mgr.triggered.connect(shift_mgr.perform)` -- both source and target are the same object. After split, source is HotkeyListener and target is ShiftSnapRestore.
**How to avoid:** Update MainWindow wiring to `hotkey_listener.triggered.connect(shift_mgr.perform)`. Verify via manual test that triple-press snap still works.
**Warning signs:** Snap key presses are detected (HotkeyListener logs trigger) but no window movement occurs.

### Pitfall 2: SnapService Delegation Chain Break
**What goes wrong:** SnapService delegates to ShiftSnapRestore methods like `update_binding()`. After splitting keyboard logic into HotkeyListener, some SnapService methods need to delegate to HotkeyListener instead, but the reference isn't set up.
**Why it happens:** SnapService holds `self._mgr` (ShiftSnapRestore), but `update_binding()` and `update_press_limit()` now belong to HotkeyListener.
**How to avoid:** Either (a) give SnapService a second reference to HotkeyListener, or (b) have ShiftSnapRestore hold a reference to HotkeyListener and delegate internally. Option (b) is cleaner -- keeps SnapService unchanged.
**Warning signs:** Changing snap key in settings has no effect.

### Pitfall 3: ExplorerService Doesn't Receive Setting Changes
**What goes wrong:** The bridge's `_apply_side_effects()` currently calls `mw._update_explorer_autosize_thread()` when `ex_auto_size` changes. After extraction, this method no longer exists on MainWindow.
**Why it happens:** Forgot to update bridge.py or MainWindow to delegate to ExplorerService.
**How to avoid:** Either (a) MainWindow keeps a thin `_update_explorer_autosize_thread()` that delegates to ExplorerService, or (b) update `bridge._apply_side_effects()` to call ExplorerService directly. Option (a) is less invasive.
**Warning signs:** Toggling explorer auto-size in settings has no effect on the background worker.

### Pitfall 4: MainWindow.closeEvent Doesn't Stop ExplorerService
**What goes wrong:** After extraction, `_stop_background_threads()` still calls `self._stop_explorer_worker()` which no longer exists on MainWindow.
**Why it happens:** Forgot to update the cleanup path in MainWindow.
**How to avoid:** Update `_stop_background_threads()` to call `self._explorer_service.stop()`.
**Warning signs:** Worker thread continues running after app close; COM errors in log.

### Pitfall 5: Test Import Errors Due to Missing Stubs
**What goes wrong:** New test scenarios for restore (D-14) need to mock `win32gui.ShowWindow(hwnd, SW_MAXIMIZE)`. If the conftest stub doesn't provide the right mock, tests fail on import.
**Why it happens:** The conftest stubs `win32gui.ShowWindow` as a basic MagicMock, but the test may need to assert it was called with specific arguments.
**How to avoid:** MagicMock already records call args. Tests can use `win32gui.ShowWindow.assert_called_with(hwnd, win32con.SW_MAXIMIZE)` with the existing conftest stubs. No new stubs needed.
**Warning signs:** `AttributeError` or `ModuleNotFoundError` when running new tests.

## Code Examples

### Negative-Coordinate Snap Position Test (D-12)
```python
# Source: verified against calculate_snap_position in virelo/services/snap.py
def test_calculate_snap_position_negative_coords():
    """Snap geometry on monitor with negative origin (left-of-primary)."""
    # Monitor at x=-1920, y=0, size 1920x1080
    x, y, w, h = calculate_snap_position(-1920, 0, 1920, 1080, 76, 76)
    assert w == 1920 * 76 // 100  # 1459
    assert h == 1080 * 76 // 100  # 820
    # x should be within the negative-coordinate monitor
    assert x == -1920 + (1920 - w) // 2  # -1690
    assert y == (1080 - h) // 2  # 130
```

### Vertical-Layout Snap Position Test (D-13)
```python
# Source: verified against calculate_snap_position in virelo/services/snap.py
def test_calculate_snap_position_vertical_layout():
    """Snap on monitor below primary (y offset)."""
    # Monitor at x=0, y=1080, size 2560x1440
    x, y, w, h = calculate_snap_position(0, 1080, 2560, 1440, 76, 76)
    assert w == 2560 * 76 // 100  # 1945
    assert h == 1440 * 76 // 100  # 1094
    assert x == (2560 - w) // 2  # 307
    assert y == 1080 + (1440 - h) // 2  # 1253
```

### Maximized-Restore Test (D-14)
```python
# Source: verified against _restore logic in virelo/services/snap.py lines 316-362
# Note: This test requires mocking win32gui and USER32 -- uses existing conftest stubs
def test_restore_maximized_window():
    """Restore of a previously-maximized window calls SW_MAXIMIZE."""
    import win32con
    import win32gui

    from virelo.services.snap import ShiftSnapRestore
    # After split, ShiftSnapRestore no longer owns keyboard logic
    # but _restore() logic is unchanged

    # Setup: simulate a window that was maximized before snap
    mgr = ShiftSnapRestore.__new__(ShiftSnapRestore)
    mgr.settings = MockSettings()
    mgr._orig_sizes = {
        12345: {
            "rect": (100, 100, 800, 600),
            "maximized": True,
        }
    }

    # Mock: need GetForegroundWindow, get_monitor_rect, GetWindowRect, ShowWindow
    # The test verifies that ShowWindow is called with SW_MAXIMIZE
    # when was_maximized is True
```

### ExplorerService Wiring in MainWindow
```python
# Source: based on MainWindow.__init__ in window.py and SnapService pattern in snap.py
# In MainWindow.__init__:
from virelo.services.explorer_service import ExplorerService

self._explorer_service = ExplorerService(self.settings, parent=self)

# Replace self._update_explorer_enabled_state() body:
def _update_explorer_enabled_state(self):
    if bool(self.settings.ex_auto_size):
        self._explorer_service.start()
    else:
        self._explorer_service.stop()

# Replace self._stop_explorer_worker() call in _stop_background_threads():
def _stop_background_threads(self):
    self._stop_capture_worker()
    self._explorer_service.stop()
    self._stop_theme_sync()
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Monolithic main.py with all logic | Package structure with services/ layer | Phase 3 (2026-04-24) | Snap already in services/snap.py |
| Explorer orchestration in MainWindow | (This phase) ExplorerService facade | Phase 4 | Decouples lifecycle from UI window |
| Snap combines hotkey + movement | (This phase) Separated concerns | Phase 4 | Testability, single responsibility |

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | HotkeyListener as QObject subclass is cleaner than plain Python class | Architecture Patterns, Pattern 1 | Low -- either approach works; QObject provides free thread-safe signal emission |
| A2 | ExplorerService should live in its own file (explorer_service.py) rather than being added to explorer_columns.py | Architecture Patterns, Pattern 2 | Low -- explorer_columns.py is already 500+ lines (Phase 3 D-04); adding lifecycle management would grow it further |
| A3 | ShiftSnapRestore holding a reference to HotkeyListener (for delegation) is cleaner than giving SnapService two references | Common Pitfalls, Pitfall 2 | Low -- implementation detail, easy to change |

## Open Questions

1. **Should SnapService expose the HotkeyListener for key rebinding?**
   - What we know: SnapService currently delegates update_binding/update_restore_key/update_press_limit to ShiftSnapRestore. After split, these methods belong to HotkeyListener.
   - What's unclear: Should SnapService hold both references, or should ShiftSnapRestore proxy to HotkeyListener?
   - Recommendation: Have SnapService hold a reference to HotkeyListener for update_binding/update_restore_key/update_press_limit, and keep the ShiftSnapRestore reference for test_snap/perform. This keeps the delegation clean without leaking HotkeyListener publicly. Alternatively, give ShiftSnapRestore a HotkeyListener reference and proxy through it.

2. **Should the `_on_press` method check `enable_snap` or should that check live higher?**
   - What we know: Currently `_on_press` checks `self.settings.enable_snap` [VERIFIED: snap.py line 199]. After split, this check would be in HotkeyListener.
   - What's unclear: Is it cleaner for HotkeyListener to check enable_snap, or should HotkeyListener always emit and let ShiftSnapRestore.perform check?
   - Recommendation: Keep the check in HotkeyListener._on_press. This avoids unnecessary signal emission and matches the current behavior exactly. [ASSUMED]

## Project Constraints (from CLAUDE.md)

Directives that affect this phase:

1. **Never reintroduce "Windows Toolbox" or "Toolbox"** -- no risk in this phase (no new strings).
2. **Never add fake UI controls** -- EXPL-01 is already satisfied; no frontend changes.
3. **Never commit generated artifacts** -- no build artifacts in scope.
4. **Never hardcode version strings** -- not relevant to this phase.
5. **Naming: Python snake_case, PascalCase classes** -- HotkeyListener, ExplorerService follow this.
6. **COM constraint from Phase 3 D-07** -- ExplorerService must NOT touch COM internals.

## Sources

### Primary (HIGH confidence)
- `virelo/services/snap.py` -- Full source read, all line numbers verified
- `virelo/app/window.py` -- Full source read, explorer lifecycle methods identified
- `virelo/workers/explorer.py` -- Full source read, COM threading pattern verified
- `virelo/bridge/bridge.py` -- Full source read, _apply_side_effects identified
- `tests/unit/test_snap_geometry.py` -- Full source read, existing test coverage verified
- `tests/conftest.py` -- Full source read, stub infrastructure verified
- `virelo/platform/win32_helpers.py` -- Full source read, helper functions verified
- `virelo/app/config.py` -- Full source read, DEFAULTS and normalize functions verified
- `.planning/phases/04-snap-and-explorer-hardening/04-CONTEXT.md` -- All decisions locked
- `.planning/phases/03-structure-and-quality/03-CONTEXT.md` -- COM constraint D-07, test tiering D-08

### Secondary (MEDIUM confidence)
- `keyboard` library API (on_press_key, on_release_key, is_pressed, unhook) -- verified from usage in snap.py, not from library docs

### Tertiary (LOW confidence)
- None

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- no new dependencies, all existing libraries verified
- Architecture: HIGH -- all patterns derived from existing codebase, all decisions locked
- Pitfalls: HIGH -- identified from concrete code analysis of existing wiring

**Research date:** 2026-04-24
**Valid until:** 2026-05-24 (stable codebase, no external dependency changes)
