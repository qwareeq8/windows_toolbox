"""Explorer column auto-size engine and worker.

Contains the tab-aware autosize engine (ExplorerAutosizeEngine),
per-tab state (TabAutosizeState), and the background QObject worker
(ExplorerAutosizeWorker).

COM Threading Constraint (D-07):
    ExplorerAutosizeWorker.run() COM initialization, Shell.Application
    caching, iter_tabs, and all COM-dependent closures MUST stay in this
    single file.  Do NOT separate COM init from COM usage.
"""

import logging
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, Protocol, TypedDict

from virelo.platform.paths import canonicalize_path

LOG = logging.getLogger("Virelo")


class _ServiceAutosize(Protocol):
    """Describe the service callback that accepts a retained COM identity."""

    def __call__(
        self,
        top_hwnd: int,
        target_path: str | None = None,
        caller_owns_com: bool = False,
        tab_identity: object | None = None,
    ) -> tuple[bool, str]: ...


_EngineAutosize = Callable[[int, int, str | None], tuple[bool, str, bool]]


class _ShellState(TypedDict):
    """Hold the worker-owned Shell.Application proxy and retry deadline."""

    app: Any | None
    retry_at: float


@dataclass
class _ComIdentityEntry:
    """Keep one worker-STA COM identity and its latest dispatch proxy alive."""

    tab_id: int
    identity: Any
    dispatch: Any


class _ComIdentityRegistry:
    """Assign stable local IDs by IUnknown equality inside one worker STA.

    PyIUnknown wrappers are unhashable, so identity comparison deliberately
    uses COM equality instead of Python object identity or enumeration order.
    Entries are pruned only after a complete Shell.Windows poll.
    """

    def __init__(self) -> None:
        self._entries: list[_ComIdentityEntry] = []
        self._next_tab_id = 1
        self._seen: set[int] | None = None

    def begin_poll(self) -> None:
        """Begin tracking identities observed by one enumeration poll."""
        self._seen = set()

    def identify(self, identity: Any, dispatch: Any) -> int:
        """Return a stable worker-local ID for one IUnknown identity."""
        if self._seen is None:
            raise RuntimeError("COM identity lookup occurred outside an active poll.")
        for entry in self._entries:
            try:
                matches = bool(identity == entry.identity)
            except Exception:
                matches = False
            if matches:
                entry.dispatch = dispatch
                self._seen.add(entry.tab_id)
                return entry.tab_id

        tab_id = self._next_tab_id
        self._next_tab_id += 1
        self._entries.append(_ComIdentityEntry(tab_id, identity, dispatch))
        self._seen.add(tab_id)
        return tab_id

    def identity_for(self, tab_id: int) -> Any | None:
        """Return the retained IUnknown wrapper for a worker-local tab ID."""
        for entry in self._entries:
            if entry.tab_id == tab_id:
                return entry.identity
        return None

    def finish_poll(self, *, complete: bool) -> None:
        """Finish a poll, pruning unseen identities only when it was complete."""
        seen = self._seen
        self._seen = None
        if complete and seen is not None:
            self._entries = [entry for entry in self._entries if entry.tab_id in seen]

    def clear(self) -> None:
        """Release every worker-owned COM reference before invalidation or exit."""
        self._entries.clear()
        self._seen = None


# Rate limiting constants
MIN_AUTOSIZE_INTERVAL_PER_TAB_MS = 200  # Minimum ms between autosize attempts per tab
GLOBAL_AUTOSIZE_RATE_LIMIT_PER_SEC = 10  # Max global autosize attempts per second
SETTLE_DELAY_MS = 150  # Time path must be stable before autosize
DEBOUNCE_DELAY_MS = 50  # Initial delay before first autosize attempt after navigation
NONINTERACTIVE_RETRY_SECONDS = 0.5
FVM_DETAILS = 4

# Circuit breaker constants
CIRCUIT_BREAKER_THRESHOLD = 5  # Failures before circuit opens
CIRCUIT_BREAKER_COOLDOWN_MS = 5000  # Cooldown after circuit opens
BACKOFF_BASE_MS = 200  # Base backoff time
BACKOFF_MAX_MS = 3000  # Maximum backoff time

# Deduplication TTL
DEDUPE_TTL_SECONDS = 300  # 5 minutes TTL for deduplication entries


@dataclass
class DedupeKey:
    """Key for deduplication: (tab_id, canonical_path, view_mode)."""

    tab_id: int
    path: str
    view_mode: int | None

    def __hash__(self):
        return hash((self.tab_id, self.path, self.view_mode))

    def __eq__(self, other):
        if not isinstance(other, DedupeKey):
            return False
        return (
            self.tab_id == other.tab_id
            and self.path == other.path
            and self.view_mode == other.view_mode
        )


@dataclass
class DedupeEntry:
    """Entry in deduplication cache with TTL."""

    key: DedupeKey
    timestamp: float
    success: bool = True


@dataclass
class TabAutosizeState:
    """
    Per-tab autosize state.

    Tracks navigation, timing, and circuit breaker state for a single tab.
    """

    tab_id: int
    hwnd: int
    path: str = ""
    view_mode: int | None = None
    navigation_token: int = 0

    # Timing
    first_seen_at: float = 0.0  # When we first saw this path
    path_stable_since: float = 0.0  # When path became stable (debounce)
    last_autosize_attempt: float = 0.0
    last_autosize_success: float = 0.0

    # Retry state
    pending_retry: bool = False
    retry_attempt: int = 0
    next_retry_at: float = 0.0

    # Circuit breaker
    consecutive_failures: int = 0
    circuit_open_until: float = 0.0
    last_error_transient: bool = False


def _try_get_window_text(hwnd):
    if hwnd is None:
        return ""
    try:
        import win32gui
    except Exception:
        return ""
    try:
        text = win32gui.GetWindowText(int(hwnd))
        return (text or "").strip()
    except Exception:
        return ""


def _resolve_explorer_path(window):
    path = ""
    try:
        loc = getattr(window, "LocationURL", None)
        if isinstance(loc, str):
            path = loc
    except Exception:
        pass
    if not path:
        try:
            doc = getattr(window, "Document", None)
            if doc is not None:
                path = str(doc.Folder.Self.Path)
        except Exception:
            pass
    if not path:
        try:
            name = getattr(window, "LocationName", None)
            if isinstance(name, str):
                path = name
        except Exception:
            pass
    if not path:
        path = _try_get_window_text(getattr(window, "HWND", None))
    return (path or "").strip()


class ExplorerAutosizeEngine:
    """
    Tab-aware engine for auto-sizing Explorer columns on navigation.

    Key features (Windows 11 tab support):
    - Tracks tabs by COM identity (tab_id), not just HWND
    - Debounce and settle detection before autosize
    - Per-tab circuit breaker with exponential backoff
    - Rate limiting (per-tab and global)
    - Scoped deduplication with TTL: (tab_id, path, view_mode)
    - Graceful handling of transient COM errors during navigation

    COM Threading Model:
    - Assumes caller (worker thread) owns COM lifetime (STA apartment)
    - Does NOT initialize/uninitialize COM per call
    - Autosize functions are called with caller_owns_com=True

    State model:
    - Per-tab state: TabAutosizeState tracking navigation, timing, failures
    - Global: deduplication cache with TTL, rate limiting
    - Bounded LRU cache for autosized paths (legacy compatibility)
    """

    # Default retry schedule: debounce, then 100ms, 250ms, 500ms, 1s
    DEFAULT_SCHEDULE = (0.05, 0.1, 0.25, 0.5, 1.0)

    # Maximum number of paths to track in autosized_paths (LRU eviction)
    MAX_AUTOSIZED_PATHS = 500

    def __init__(
        self,
        iter_tabs: Callable[[], list[tuple[Any, ...]]],
        autosize_try: _EngineAutosize,
        autosize_full: _EngineAutosize,
        is_window_interactive: Callable[[int], bool],
        schedule: tuple[float, ...] | None = None,
        persist_paths: bool = False,
    ):
        """
        Initialize the tab-aware autosize engine.

        Args:
            iter_tabs: Callable returning list of (hwnd, tab_id, path, view_mode) tuples
            autosize_try: Quick autosize attempt (hwnd, tab_id, path),
                returns (success, method, is_transient)
            autosize_full: Full autosize with all fallbacks (hwnd, tab_id, path),
                returns (success, method, is_transient)
            is_window_interactive: Check if window is visible and ready
            schedule: Retry schedule as tuple of delays in seconds
            persist_paths: If True, persist autosized paths (not recommended)
        """
        from collections import OrderedDict

        self._iter_tabs = iter_tabs
        self._autosize_try = autosize_try
        self._autosize_full = autosize_full
        self._is_window_interactive = is_window_interactive
        self._schedule = schedule if schedule else self.DEFAULT_SCHEDULE
        self._persist_paths = persist_paths

        # Per-tab state keyed by the supplied COM identity. HWND remains part
        # of the key as a defensive namespace if an identity is ever reused.
        self.tab_state: dict[tuple[int, int], TabAutosizeState] = {}

        # Deduplication cache with TTL: {DedupeKey: DedupeEntry}
        self._dedupe_cache: dict[DedupeKey, DedupeEntry] = {}

        # Global rate limiting
        self._global_autosize_times: list[float] = []  # Recent autosize timestamps

        # Navigation token counter (increments on each navigation)
        self._token_counter = 0

        # Legacy compatibility: bounded LRU cache for autosized paths
        # OrderedDict provides O(1) move_to_end for LRU behavior
        self.autosized_paths: OrderedDict[str, float] = OrderedDict()

        # Legacy compatibility: window_state maps hwnd to primary tab
        self.window_state: dict[int, dict] = {}
        self.pending: dict[int, dict] = {}

        # Idle backoff: poll fast only while the Explorer window set is
        # actually changing; a resident tray app must not enumerate COM
        # windows 6+ times per second around the clock.
        self._last_activity = 0.0
        self._prev_live_signatures: set = set()

    def _next_token(self) -> int:
        """Generate the next navigation token."""
        self._token_counter += 1
        return self._token_counter

    def _is_dedupe_valid(self, key: DedupeKey, now: float) -> bool:
        """Check if a deduplication entry is still valid (within TTL)."""
        entry = self._dedupe_cache.get(key)
        if entry is None:
            return False
        if now - entry.timestamp > DEDUPE_TTL_SECONDS:
            # TTL expired
            del self._dedupe_cache[key]
            return False
        return entry.success

    def _record_dedupe(self, key: DedupeKey, now: float, success: bool):
        """Record a deduplication entry."""
        self._dedupe_cache[key] = DedupeEntry(key=key, timestamp=now, success=success)

    def _clean_dedupe_cache(self, now: float):
        """Remove expired entries from deduplication cache."""
        expired = [
            k for k, v in self._dedupe_cache.items() if now - v.timestamp > DEDUPE_TTL_SECONDS
        ]
        for k in expired:
            del self._dedupe_cache[k]

    def _check_global_rate_limit(self, now: float) -> bool:
        """Check if global rate limit allows an autosize attempt."""
        # Clean old entries
        cutoff = now - 1.0
        self._global_autosize_times = [t for t in self._global_autosize_times if t > cutoff]
        return len(self._global_autosize_times) < GLOBAL_AUTOSIZE_RATE_LIMIT_PER_SEC

    def _record_global_autosize(self, now: float):
        """Record a global autosize attempt."""
        self._global_autosize_times.append(now)

    def _calculate_backoff(self, failures: int) -> float:
        """Calculate exponential backoff time in seconds."""
        if failures <= 0:
            return 0
        backoff_ms = min(BACKOFF_BASE_MS * (2 ** (failures - 1)), BACKOFF_MAX_MS)
        return backoff_ms / 1000.0

    def _record_autosized_path(self, path: str, timestamp: float):
        """
        Record a path as autosized with LRU eviction.

        Args:
            path: Canonical path that was autosized
            timestamp: Time of autosizing
        """
        if path in self.autosized_paths:
            # Move to end (mark as recently used)
            self.autosized_paths.move_to_end(path)
            self.autosized_paths[path] = timestamp
        else:
            self.autosized_paths[path] = timestamp
            # Evict oldest if over limit
            if len(self.autosized_paths) > self.MAX_AUTOSIZED_PATHS:
                self.autosized_paths.popitem(last=False)

    def step(self, now: float) -> float:
        """
        Process one step of the tab-aware autosize engine.

        Args:
            now: Current monotonic timestamp.

        Returns:
            Recommended delay before next step
        """
        log = LOG

        # Periodically clean dedupe cache
        self._clean_dedupe_cache(now)

        try:
            tabs = list(self._iter_tabs() or [])
        except Exception as e:
            log.warning("ExplorerAutosizeEngine.step: iter_tabs failed: %s", e)
            tabs = []

        if tabs:
            log.debug("ExplorerAutosizeEngine.step: found %d Explorer tabs", len(tabs))

        live_tab_keys: set[tuple[int, int]] = set()
        live_tab_signatures: set[tuple[int, int, str, int | None]] = set()

        for tab_info in tabs:
            # Unpack tab info: (hwnd, tab_id, path, view_mode, [dispatch])
            supplied_tab_id: Any = None
            if len(tab_info) >= 5:
                hwnd, supplied_tab_id, raw_path, view_mode = tab_info[:4]
            elif len(tab_info) >= 4:
                hwnd, supplied_tab_id, raw_path, view_mode = tab_info[:4]
            elif len(tab_info) >= 3:
                hwnd, supplied_tab_id, raw_path = tab_info[:3]
                view_mode = None
            elif len(tab_info) >= 2:
                # Legacy format: (hwnd, path)
                hwnd, raw_path = tab_info[:2]
                view_mode = None
            else:
                continue

            raw_path = "" if raw_path is None else str(raw_path)
            canonical_path = canonicalize_path(raw_path)
            try:
                tab_id = int(supplied_tab_id)
            except (TypeError, ValueError):
                # Legacy callers do not provide a COM identity. Keep their
                # previous path-based behavior without weakening modern tabs.
                tab_id = hash((int(hwnd), canonical_path))

            tab_key = (int(hwnd), tab_id)
            live_tab_keys.add(tab_key)
            live_tab_signatures.add((int(hwnd), tab_id, canonical_path, view_mode))

            state = self.tab_state.get(tab_key)

            if state is None:
                state = TabAutosizeState(
                    tab_id=tab_id,
                    hwnd=int(hwnd),
                    path=canonical_path,
                    view_mode=view_mode,
                    navigation_token=self._next_token(),
                    first_seen_at=now,
                    path_stable_since=now,
                )
                self.tab_state[tab_key] = state
                log.debug(
                    "ExplorerAutosizeEngine.step: new tab hwnd=%s tab_id=%s path='%s'",
                    hwnd,
                    tab_id,
                    canonical_path,
                )

                if view_mode in (None, FVM_DETAILS):
                    debounce_delay = DEBOUNCE_DELAY_MS / 1000.0
                    state.next_retry_at = now + debounce_delay
                    state.pending_retry = True
                    log.debug(
                        "ExplorerAutosizeEngine.step: scheduled debounced autosize "
                        "in %.3fs for hwnd=%s tab_id=%s path='%s'",
                        debounce_delay,
                        hwnd,
                        tab_id,
                        canonical_path,
                    )
                continue  # Don't attempt autosize this iteration

            if canonical_path != state.path:
                old_path = state.path
                # A successful fit is scoped to one navigation, not to the
                # path for the full deduplication TTL. Remove prior results for
                # this destination so A -> B -> A triggers a fresh fit.
                prior_destination_keys = [
                    dedupe_entry_key
                    for dedupe_entry_key in self._dedupe_cache
                    if dedupe_entry_key.tab_id == state.tab_id
                    and dedupe_entry_key.path == canonical_path
                ]
                for dedupe_entry_key in prior_destination_keys:
                    self._dedupe_cache.pop(dedupe_entry_key, None)
                state.path = canonical_path
                state.view_mode = view_mode
                state.navigation_token = self._next_token()
                state.first_seen_at = now
                state.path_stable_since = now
                state.last_autosize_attempt = 0.0
                state.retry_attempt = 0
                state.consecutive_failures = 0
                state.circuit_open_until = 0.0
                state.pending_retry = view_mode in (None, FVM_DETAILS)
                state.next_retry_at = now + (DEBOUNCE_DELAY_MS / 1000.0)
                log.debug(
                    "ExplorerAutosizeEngine.step: tab navigated hwnd=%s tab_id=%s '%s' -> '%s'",
                    hwnd,
                    tab_id,
                    old_path,
                    canonical_path,
                )
                continue

            if view_mode != state.view_mode:
                old_view_mode = state.view_mode
                log.debug(
                    "ExplorerAutosizeEngine.step: view mode changed "
                    "for hwnd=%s tab_id=%s path='%s': %s -> %s",
                    hwnd,
                    tab_id,
                    canonical_path,
                    old_view_mode,
                    view_mode,
                )
                state.view_mode = view_mode
                state.path_stable_since = now
                if view_mode == FVM_DETAILS:
                    # A previous Details-view result must not suppress a fresh
                    # fit after the user leaves and returns to Details.
                    self._dedupe_cache.pop(
                        DedupeKey(
                            tab_id=state.tab_id,
                            path=canonical_path,
                            view_mode=FVM_DETAILS,
                        ),
                        None,
                    )
                    state.pending_retry = True
                    state.retry_attempt = 0
                    state.consecutive_failures = 0
                    state.circuit_open_until = 0.0
                    state.next_retry_at = now + (DEBOUNCE_DELAY_MS / 1000.0)
                elif view_mode is not None:
                    state.pending_retry = False
                continue

            # Re-arm a tab whose circuit-breaker cooldown has expired so the
            # cooldown actually leads to one more retry instead of parking the
            # tab forever.
            if (
                not state.pending_retry
                and state.circuit_open_until > 0
                and now >= state.circuit_open_until
            ):
                state.pending_retry = True
                state.retry_attempt = 0
                state.consecutive_failures = 0
                state.circuit_open_until = 0.0
                state.next_retry_at = now
                state.path_stable_since = now
                log.debug(
                    "ExplorerAutosizeEngine.step: circuit cooldown elapsed, re-arming hwnd=%s",
                    hwnd,
                )

            # Check if we should attempt autosize for this entry
            if not state.pending_retry:
                continue

            # Check circuit breaker
            if state.circuit_open_until > now:
                log.debug(
                    "ExplorerAutosizeEngine.step: circuit open for hwnd=%s until %.3f",
                    hwnd,
                    state.circuit_open_until - now,
                )
                state.next_retry_at = state.circuit_open_until
                continue

            # Check if debounce/retry timer has elapsed
            if now < state.next_retry_at:
                continue

            # Check settle requirement: path must be stable for SETTLE_DELAY_MS
            settle_required = SETTLE_DELAY_MS / 1000.0
            if now - state.path_stable_since < settle_required:
                log.debug(
                    "ExplorerAutosizeEngine.step: waiting for settle, hwnd=%s path='%s'",
                    hwnd,
                    canonical_path,
                )
                state.next_retry_at = state.path_stable_since + settle_required
                continue

            # Check per-tab rate limit
            per_tab_interval = MIN_AUTOSIZE_INTERVAL_PER_TAB_MS / 1000.0
            if now - state.last_autosize_attempt < per_tab_interval:
                log.debug(
                    "ExplorerAutosizeEngine.step: rate limited for hwnd=%s path='%s'",
                    hwnd,
                    canonical_path,
                )
                state.next_retry_at = state.last_autosize_attempt + per_tab_interval
                continue

            # Check global rate limit
            if not self._check_global_rate_limit(now):
                log.debug("ExplorerAutosizeEngine.step: global rate limit reached")
                oldest_attempt = min(self._global_autosize_times, default=now)
                state.next_retry_at = max(state.next_retry_at, oldest_attempt + 1.0)
                continue

            dedupe_key = DedupeKey(
                tab_id=state.tab_id,
                path=canonical_path,
                view_mode=view_mode,
            )
            if self._is_dedupe_valid(dedupe_key, now):
                log.debug(
                    "ExplorerAutosizeEngine.step: already autosized "
                    "(dedupe hit) hwnd=%s path='%s' view=%s",
                    hwnd,
                    canonical_path,
                    view_mode,
                )
                state.pending_retry = False
                continue

            # Check if window is interactive
            if not self._is_window_interactive(hwnd):
                log.debug("ExplorerAutosizeEngine.step: hwnd=%s not interactive, skipping", hwnd)
                state.next_retry_at = now + NONINTERACTIVE_RETRY_SECONDS
                continue

            # Attempt autosize
            log.debug(
                "ExplorerAutosizeEngine.step: attempting autosize for hwnd=%s path='%s' attempt=%d",
                hwnd,
                canonical_path,
                state.retry_attempt,
            )

            state.last_autosize_attempt = now
            self._record_global_autosize(now)

            ok = False
            method = "none"
            is_transient = False

            try:
                # Use quick autosize for first few attempts, then full
                # Pass the stable tab identity. A path is insufficient because
                # two tabs under one HWND may show the same folder.
                if state.retry_attempt <= 2:
                    result = self._autosize_try(hwnd, state.tab_id, canonical_path)
                else:
                    result = self._autosize_full(hwnd, state.tab_id, canonical_path)

                if isinstance(result, tuple):
                    if len(result) >= 3:
                        ok, method, is_transient = result[:3]
                    elif len(result) >= 2:
                        ok, method = result[:2]
                        is_transient = False
                    else:
                        ok = bool(result[0])
                        method = "unknown" if ok else "none"
                        is_transient = False
                else:
                    ok = bool(result)
                    method = "unknown" if ok else "none"
                    is_transient = False

                log.debug(
                    "ExplorerAutosizeEngine.step: autosize result: "
                    "ok=%s method=%s transient=%s hwnd=%s",
                    ok,
                    method,
                    is_transient,
                    hwnd,
                )
            except Exception as e:
                ok = False
                method = "error"
                is_transient = True  # Assume exceptions are transient
                log.warning(
                    "ExplorerAutosizeEngine.step: autosize exception for hwnd=%s: %s", hwnd, e
                )

            if ok:
                # Success!
                state.pending_retry = False
                state.consecutive_failures = 0
                state.last_autosize_success = now
                self._record_dedupe(dedupe_key, now, success=True)

                # Legacy compatibility: record path with LRU eviction
                self._record_autosized_path(canonical_path, now)

                log.debug(
                    "ExplorerAutosizeEngine.step: SUCCESS - autosized hwnd=%s path='%s' method=%s",
                    hwnd,
                    canonical_path,
                    method,
                )
            else:
                # Failure
                state.retry_attempt += 1
                state.last_error_transient = is_transient

                if is_transient:
                    # Transient error - schedule retry with backoff
                    state.consecutive_failures += 1
                    backoff = self._calculate_backoff(state.consecutive_failures)

                    if state.retry_attempt < len(self._schedule):
                        retry_delay = max(self._schedule[state.retry_attempt], backoff)
                        state.next_retry_at = now + retry_delay
                        log.debug(
                            "ExplorerAutosizeEngine.step: transient failure, "
                            "retry in %.3fs for hwnd=%s",
                            retry_delay,
                            hwnd,
                        )
                    else:
                        # Check circuit breaker
                        if state.consecutive_failures >= CIRCUIT_BREAKER_THRESHOLD:
                            cooldown = CIRCUIT_BREAKER_COOLDOWN_MS / 1000.0
                            state.circuit_open_until = now + cooldown
                            state.pending_retry = False
                            log.warning(
                                "ExplorerAutosizeEngine.step: circuit breaker "
                                "opened for hwnd=%s, cooldown %.1fs",
                                hwnd,
                                cooldown,
                            )
                        else:
                            state.pending_retry = False
                            log.warning(
                                "ExplorerAutosizeEngine.step: exhausted retries "
                                "for hwnd=%s tab_id=%s",
                                hwnd,
                                state.tab_id,
                            )
                else:
                    # Non-transient error (e.g., not in Details view) - don't retry
                    state.pending_retry = False
                    self._record_dedupe(dedupe_key, now, success=False)
                    log.info(
                        "ExplorerAutosizeEngine.step: non-transient failure, not retrying hwnd=%s",
                        hwnd,
                    )

        # Clean up tab identities that no longer exist.
        closed_entries = [tab_key for tab_key in self.tab_state if tab_key not in live_tab_keys]
        for closed_tab_key in closed_entries:
            del self.tab_state[closed_tab_key]
            log.debug("ExplorerAutosizeEngine.step: removed closed entry %s", closed_tab_key)

        # Update legacy window_state for compatibility
        self.window_state = {
            state.hwnd: {"path": state.path, "token": state.navigation_token}
            for state in self.tab_state.values()
        }
        self.pending = {
            state.hwnd: {
                "token": state.navigation_token,
                "path": state.path,
                "attempt": state.retry_attempt,
                "next": state.next_retry_at,
            }
            for state in self.tab_state.values()
            if state.pending_retry
        }

        # Record activity whenever the set of open tabs, paths, or views changed.
        # Startup counts as activity so a freshly enabled worker responds fast.
        if self._last_activity == 0.0 or live_tab_signatures != self._prev_live_signatures:
            self._last_activity = now
            self._prev_live_signatures = set(live_tab_signatures)

        return self._next_delay(now)

    def _next_delay(self, now: float) -> float:
        """Calculate the delay until the next step should run."""
        # Check if any tabs have pending work
        pending_tabs = [s for s in self.tab_state.values() if s.pending_retry]

        if not pending_tabs:
            # No pending work: decay the poll rate with inactivity. The cost
            # of a poll is a full cross-process COM enumeration, so idle
            # cadence matters more than first-detection latency.
            idle = now - self._last_activity
            if idle < 5.0:
                return 0.15
            if idle < 30.0:
                return 0.5
            return 1.0

        # Find the soonest pending retry
        next_due = min(s.next_retry_at for s in pending_tabs)
        delay = next_due - now

        # Respect future retry deadlines instead of repeatedly enumerating COM
        # while a minimized window is waiting to become interactive. A 50 ms
        # floor also prevents a past-due retry from creating a 200 Hz loop.
        return max(0.05, min(0.5, delay))

    def clear_path_history(self):
        """Clear the set of autosized paths and deduplication cache (for testing or reset)."""
        self.autosized_paths.clear()
        self._dedupe_cache.clear()
        for state in self.tab_state.values():
            state.consecutive_failures = 0
            state.circuit_open_until = 0.0

    @property
    def last_paths(self) -> dict[int, str]:
        """Compatibility property: return current paths per window."""
        return {hwnd: st.get("path", "") for hwnd, st in self.window_state.items()}


# Conditional PySide6 import for CI compatibility (D-10)
QtCore: Any
try:
    from PySide6 import QtCore as _QtCore

    QtCore = _QtCore
except Exception:  # pragma: no cover - PySide6 unavailable in some test envs
    QtCore = None


if QtCore is not None:

    class ExplorerAutosizeWorker(QtCore.QObject):
        """
        Background worker for Explorer column auto-sizing.

        Runs in a separate thread and monitors Explorer windows/tabs for navigation,
        auto-sizing columns on the first visit to each path.

        Windows 11 tab support:
        - Tracks tabs by COM identity (tab_id), not just HWND
        - Properly handles multiple tabs in a single window
        - Includes debounce, settle detection, rate limiting, and circuit breakers
        """

        finished = QtCore.Signal()

        # Default retry schedule: debounce, then 100ms, 250ms, 500ms, 1s
        DEFAULT_SCHEDULE = (0.05, 0.1, 0.25, 0.5, 1.0)

        def __init__(
            self,
            autosize_try: _ServiceAutosize,
            autosize_full: _ServiceAutosize,
            is_window_interactive: Callable[[int], bool],
            schedule: tuple[float, ...] | None = None,
        ):
            """
            Initialize the autosize worker.

            Args:
                autosize_try: Quick autosize attempt, returns
                    (success, method) or (success, method, is_transient)
                autosize_full: Full autosize with all fallbacks,
                    returns (success, method) or
                    (success, method, is_transient)
                is_window_interactive: Check if window is visible
                schedule: Retry schedule as tuple of delays in seconds
            """
            super().__init__()
            self._autosize_try_service = autosize_try
            self._autosize_full_service = autosize_full
            self._is_window_interactive = is_window_interactive
            self._schedule = schedule if schedule else self.DEFAULT_SCHEDULE
            self._stop = threading.Event()
            # Win32 event mirrored with _stop so the worker's message wait can
            # wake immediately on stop instead of polling a Python flag.
            try:
                import win32event

                self._stop_win32_event = win32event.CreateEvent(None, True, False, None)
            except Exception:  # pragma: no cover - pywin32 absent off-Windows
                self._stop_win32_event = None

        def stop(self):
            """Signal the worker to stop."""
            self._stop.set()
            if self._stop_win32_event is not None:
                try:
                    import win32event

                    win32event.SetEvent(self._stop_win32_event)
                except Exception:
                    pass

        @QtCore.Slot()
        def run(self):
            """
            Main worker loop with proper COM lifecycle management.

            COM Threading Model:
            - Initializes COM ONCE for this thread (STA apartment)
            - Caches Shell.Application to avoid repeated creation/destruction
            - Pumps COM messages to prevent RPC disconnection errors
            - Uninitializes COM only at shutdown after releasing all objects
            """
            import logging

            import pythoncom
            import pywintypes
            import win32com.client
            import win32con
            import win32event
            import win32gui

            log = logging.getLogger("Virelo")
            log.info("Explorer autosize worker: initializing COM once per thread (STA)...")

            # COM is initialized ONCE for this worker thread (STA)
            try:
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            except Exception:
                # Fallback if CoInitializeEx is unavailable
                pythoncom.CoInitialize()
            log.info("Explorer autosize worker: COM initialized (STA)")

            # Cache Shell.Application to avoid creating/destroying it every poll
            # cycle (avoids RPC errors 0x80010108, 0x800706b5). A creation
            # failure is retried with a cooldown instead of disabling the
            # feature for the process lifetime.
            shell_state: _ShellState = {"app": None, "retry_at": 0.0}
            identity_registry = _ComIdentityRegistry()

            def invalidate_shell_app(reason: str, retry_delay: float = 1.0) -> None:
                """Release a disconnected proxy so the next poll can recreate it."""
                if shell_state["app"] is not None:
                    log.debug(
                        "Explorer autosize worker: invalidating Shell.Application: %s",
                        reason,
                    )
                identity_registry.clear()
                shell_state["app"] = None
                shell_state["retry_at"] = time.monotonic() + retry_delay

            def ensure_shell_app():
                if shell_state["app"] is not None:
                    return shell_state["app"]
                if time.monotonic() < shell_state["retry_at"]:
                    return None
                try:
                    shell_state["app"] = win32com.client.Dispatch("Shell.Application")
                    shell_state["retry_at"] = 0.0
                    log.info("Explorer autosize worker: Shell.Application cached")
                except Exception as e:
                    shell_state["retry_at"] = time.monotonic() + 30.0
                    log.warning(
                        "Explorer autosize worker: Shell.Application creation failed "
                        "(retrying in 30s): %s",
                        e,
                    )
                return shell_state["app"]

            ensure_shell_app()

            try:

                def _get_view_mode(w) -> int | None:
                    """Get view mode safely, returning None if not available."""
                    try:
                        doc = getattr(w, "Document", None)
                        if doc is None:
                            return None
                        return int(getattr(doc, "CurrentViewMode"))
                    except (pywintypes.com_error, OSError, AttributeError, Exception):
                        return None

                def iter_tabs():
                    """Enumerate Explorer tabs with stable worker-owned COM identities."""

                    def stopping() -> bool:
                        return self._stop.is_set()

                    shell_app = None if stopping() else ensure_shell_app()
                    if stopping() or shell_app is None:
                        return []

                    out: list[tuple[int, int, str, int | None]] = []
                    windows = None
                    window_list = []
                    enumeration_complete = False
                    identity_registry.begin_poll()

                    try:
                        try:
                            if stopping():
                                return out
                            windows = shell_app.Windows()
                            if windows is None:
                                invalidate_shell_app("Windows() returned None")
                                return out
                            if stopping():
                                return out
                            count = int(windows.Count)
                        except (
                            pywintypes.com_error,
                            pythoncom.com_error,
                            OSError,
                            Exception,
                        ) as exc:
                            log.debug("iter_tabs: Windows enumeration failed: %s", exc)
                            invalidate_shell_app("Windows() enumeration failed")
                            return out

                        enumeration_complete = True
                        item_failures = 0
                        for index in range(count):
                            if stopping():
                                enumeration_complete = False
                                return out
                            try:
                                window = windows.Item(index)
                            except (
                                pywintypes.com_error,
                                pythoncom.com_error,
                                OSError,
                                Exception,
                            ):
                                item_failures += 1
                                enumeration_complete = False
                                continue
                            if window is not None:
                                window_list.append(window)

                        if count > 0 and item_failures == count:
                            invalidate_shell_app("every Windows.Item call failed")
                            return out

                        log.debug(
                            "iter_tabs: Shell.Application returned %d windows",
                            len(window_list),
                        )
                        for window in window_list:
                            if stopping():
                                enumeration_complete = False
                                break

                            try:
                                hwnd = int(window.HWND or 0)
                            except (
                                pywintypes.com_error,
                                pythoncom.com_error,
                                OSError,
                                AttributeError,
                                Exception,
                            ):
                                enumeration_complete = False
                                continue
                            if not hwnd:
                                enumeration_complete = False
                                continue

                            try:
                                window_class = win32gui.GetClassName(hwnd)
                            except Exception:
                                enumeration_complete = False
                                continue
                            if window_class not in ("CabinetWClass", "ExploreWClass"):
                                continue

                            try:
                                identity = window._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
                                tab_id = identity_registry.identify(identity, window)
                            except (
                                pywintypes.com_error,
                                pythoncom.com_error,
                                OSError,
                                AttributeError,
                                Exception,
                            ) as exc:
                                enumeration_complete = False
                                log.debug(
                                    "iter_tabs: could not retain COM identity for hwnd=%s: %s",
                                    hwnd,
                                    exc,
                                )
                                continue

                            path = _resolve_explorer_path(window)
                            view_mode = _get_view_mode(window)
                            out.append((hwnd, tab_id, path, view_mode))
                            log.debug(
                                "iter_tabs: found tab hwnd=%s tab_id=%s view=%s",
                                hwnd,
                                tab_id,
                                view_mode,
                            )
                        return out
                    except (
                        pywintypes.com_error,
                        pythoncom.com_error,
                        OSError,
                        Exception,
                    ) as exc:
                        enumeration_complete = False
                        log.debug("iter_tabs: failed while processing windows: %s", exc)
                        invalidate_shell_app("Windows() processing failed")
                        return out
                    finally:
                        window_list.clear()
                        windows = None
                        identity_registry.finish_poll(
                            complete=enumeration_complete and not stopping()
                        )

                def autosize_try_wrapper(
                    hwnd: int, tab_id: int, target_path: str | None
                ) -> tuple[bool, str, bool]:
                    """Autosize the exact retained COM tab identity."""
                    if self._stop.is_set():
                        return (False, "stopped", False)
                    tab_identity = identity_registry.identity_for(tab_id)
                    if tab_identity is None:
                        return (False, "not-found", True)
                    try:
                        # Worker thread owns COM lifecycle - tell autosize not to init/uninit
                        result = self._autosize_try_service(
                            hwnd,
                            target_path=target_path,
                            caller_owns_com=True,
                            tab_identity=tab_identity,
                        )
                        if isinstance(result, tuple) and len(result) >= 2:
                            ok, method = result[:2]
                            # "not-details" is permanent until the user changes
                            # the view; retrying it five times per navigation
                            # is pure COM churn.
                            is_transient = not ok and method in ("none", "error", "not-found")
                            return (ok, method, is_transient)
                        return (bool(result), "unknown" if result else "none", True)
                    except Exception as e:
                        log.debug(
                            "autosize_try_wrapper: exception for hwnd=%s path=%s: %s",
                            hwnd,
                            target_path,
                            e,
                        )
                        return (False, "error", True)  # Exceptions are transient

                def autosize_full_wrapper(
                    hwnd: int, tab_id: int, target_path: str | None
                ) -> tuple[bool, str, bool]:
                    """Run the full strategy against the exact retained COM tab."""
                    if self._stop.is_set():
                        return (False, "stopped", False)
                    tab_identity = identity_registry.identity_for(tab_id)
                    if tab_identity is None:
                        return (False, "not-found", True)
                    try:
                        # Worker thread owns COM lifecycle - tell autosize not to init/uninit
                        result = self._autosize_full_service(
                            hwnd,
                            target_path=target_path,
                            caller_owns_com=True,
                            tab_identity=tab_identity,
                        )
                        if isinstance(result, tuple) and len(result) >= 2:
                            ok, method = result[:2]
                            is_transient = not ok and method in ("none", "error", "not-found")
                            return (ok, method, is_transient)
                        return (bool(result), "unknown" if result else "none", True)
                    except Exception as e:
                        log.debug(
                            "autosize_full_wrapper: exception for hwnd=%s path=%s: %s",
                            hwnd,
                            target_path,
                            e,
                        )
                        return (False, "error", True)

                log.info("Explorer autosize worker: started with schedule=%s", self._schedule)
                engine = ExplorerAutosizeEngine(
                    iter_tabs,
                    autosize_try_wrapper,
                    autosize_full_wrapper,
                    self._is_window_interactive,
                    schedule=self._schedule,
                )
                log.info("Explorer autosize worker: engine created, entering main loop")
                loop_count = 0

                while not self._stop.is_set():
                    # Double-check stop flag before any COM work
                    if self._stop.is_set():
                        break
                    try:
                        now = time.monotonic()
                        sleep_for = engine.step(now)
                        loop_count += 1

                        # Log metrics periodically
                        if loop_count <= 5 or loop_count % 50 == 0:
                            log.debug(
                                "Explorer autosize: loop=%d tabs=%d pending=%d "
                                "cache=%d paths=%d sleep=%.3fs",
                                loop_count,
                                len(engine.tab_state),
                                len([s for s in engine.tab_state.values() if s.pending_retry]),
                                len(engine._dedupe_cache),
                                len(engine.autosized_paths),
                                sleep_for,
                            )
                    except Exception as exc:
                        log.exception("Explorer autosize worker: step failed", exc_info=exc)
                        sleep_for = 0.2

                    # Pump COM messages to keep STA apartment responsive
                    # This prevents RPC_E_DISCONNECTED and other STA threading issues
                    try:
                        if pythoncom.PumpWaitingMessages():
                            # WM_QUIT received
                            break
                    except Exception:
                        pass

                    # Sleep without spinning: MsgWaitForMultipleObjects wakes
                    # only for the stop event, an incoming window message
                    # (pumped for the STA), or the timeout. The old loop woke
                    # 200 times per second around the clock.
                    deadline = time.monotonic() + sleep_for
                    while not self._stop.is_set():
                        remaining_ms = int((deadline - time.monotonic()) * 1000)
                        if remaining_ms <= 0:
                            break
                        if self._stop_win32_event is None:
                            time.sleep(min(0.05, remaining_ms / 1000.0))
                            continue
                        rc = win32event.MsgWaitForMultipleObjects(
                            [self._stop_win32_event],
                            False,
                            remaining_ms,
                            win32con.QS_ALLINPUT,
                        )
                        if rc == win32event.WAIT_OBJECT_0:
                            break  # Stop requested.
                        if rc == win32event.WAIT_TIMEOUT:
                            break  # Slept the full delay.
                        # rc == WAIT_OBJECT_0 + 1: window messages are pending.
                        try:
                            if pythoncom.PumpWaitingMessages():
                                self._stop.set()  # WM_QUIT
                                break
                        except Exception:
                            pass
            finally:
                # CRITICAL: Release all COM objects BEFORE uninitializing COM
                # This prevents crashes from lingering COM references
                identity_registry.clear()
                shell_state["app"] = None
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                log.info(
                    "Explorer autosize worker: exiting after %d loops.",
                    loop_count if "loop_count" in dir() else 0,
                )
                self.finished.emit()
