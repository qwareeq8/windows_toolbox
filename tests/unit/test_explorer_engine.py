"""Regression tests for Explorer tab tracking and autosize scheduling."""

from virelo.workers.explorer import ExplorerAutosizeEngine, _ComIdentityRegistry


def _make_engine(tabs, calls, *, interactive=True):
    def autosize(hwnd, tab_id, target_path):
        calls.append((hwnd, tab_id, target_path))
        return True, "com", False

    return ExplorerAutosizeEngine(
        lambda: list(tabs),
        autosize,
        autosize,
        lambda hwnd: interactive,
    )


def test_same_path_tabs_keep_independent_state():
    """Two tabs on the same path should not collapse into one state entry."""
    tabs = [
        (100, 11, r"C:\Shared", 4),
        (100, 12, r"C:\Shared", 4),
    ]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls)

    engine.step(10.0)
    engine.step(10.2)

    assert set(engine.tab_state) == {(100, 11), (100, 12)}
    assert {state.tab_id for state in engine.tab_state.values()} == {11, 12}
    assert calls == [
        (100, 11, r"c:\shared"),
        (100, 12, r"c:\shared"),
    ]
    assert {key.tab_id for key in engine._dedupe_cache} == {11, 12}


def test_navigation_updates_state_for_stable_tab_identity():
    """A path change should update one tab state and issue a new navigation token."""
    tabs = [(100, 11, r"C:\First", 4)]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls)

    engine.step(20.0)
    engine.step(20.2)
    state = engine.tab_state[(100, 11)]
    first_token = state.navigation_token

    tabs[0] = (100, 11, r"C:\Second", 4)
    engine.step(21.0)

    assert len(engine.tab_state) == 1
    assert engine.tab_state[(100, 11)] is state
    assert state.path == r"c:\second"
    assert state.navigation_token > first_token
    assert state.pending_retry is True

    engine.step(21.2)
    assert calls == [(100, 11, r"c:\first"), (100, 11, r"c:\second")]


def test_returning_to_a_path_after_navigation_fits_again():
    """Returning from B to A should not reuse A's prior navigation result."""
    tabs = [(100, 11, r"C:\First", 4)]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls)

    engine.step(25.0)
    engine.step(25.2)
    tabs[0] = (100, 11, r"C:\Second", 4)
    engine.step(26.0)
    engine.step(26.2)
    tabs[0] = (100, 11, r"C:\First", 4)
    engine.step(27.0)
    engine.step(27.2)

    assert calls == [
        (100, 11, r"c:\first"),
        (100, 11, r"c:\second"),
        (100, 11, r"c:\first"),
    ]


def test_known_non_details_view_waits_until_details():
    """A known non-Details tab should remain idle and rearm on entering Details."""
    tabs = [(100, 11, r"C:\Folder", 3)]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls)

    engine.step(30.0)
    engine.step(40.0)
    assert calls == []
    assert engine.tab_state[(100, 11)].pending_retry is False

    tabs[0] = (100, 11, r"C:\Folder", 4)
    engine.step(41.0)
    assert engine.tab_state[(100, 11)].pending_retry is True

    engine.step(41.2)
    assert calls == [(100, 11, r"c:\folder")]


def test_returning_to_details_bypasses_prior_dedupe_result():
    """Returning to Details should fit again even after an earlier success."""
    tabs = [(100, 11, r"C:\Folder", 4)]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls)

    engine.step(45.0)
    engine.step(45.2)
    tabs[0] = (100, 11, r"C:\Folder", 3)
    engine.step(46.0)
    tabs[0] = (100, 11, r"C:\Folder", 4)
    engine.step(47.0)
    engine.step(47.2)

    assert calls == [(100, 11, r"c:\folder"), (100, 11, r"c:\folder")]


def test_noninteractive_tab_uses_retry_deadline_instead_of_spinning():
    """A minimized or hidden tab should wait before another eligibility check."""
    tabs = [(100, 11, r"C:\Folder", 4)]
    calls: list[tuple[int, int, str | None]] = []
    engine = _make_engine(tabs, calls, interactive=False)

    engine.step(50.0)
    delay = engine.step(50.2)

    assert calls == []
    assert engine.tab_state[(100, 11)].next_retry_at == 50.7
    assert delay == 0.5


class _FakeIdentity:
    """Compare by COM-like object identity while using fresh wrappers."""

    def __init__(self, value):
        self.value = value

    def __eq__(self, other):
        return isinstance(other, _FakeIdentity) and self.value == other.value


def test_com_identity_registry_survives_reordering_and_prunes_complete_polls():
    """Enumeration order changes must not change stable worker-local tab IDs."""
    registry = _ComIdentityRegistry()
    registry.begin_poll()
    first_id = registry.identify(_FakeIdentity("first"), object())
    second_id = registry.identify(_FakeIdentity("second"), object())
    registry.finish_poll(complete=True)

    registry.begin_poll()
    reordered_second_id = registry.identify(_FakeIdentity("second"), object())
    reordered_first_id = registry.identify(_FakeIdentity("first"), object())
    registry.finish_poll(complete=True)

    assert (reordered_first_id, reordered_second_id) == (first_id, second_id)

    registry.begin_poll()
    registry.identify(_FakeIdentity("first"), object())
    registry.finish_poll(complete=True)
    assert registry.identity_for(second_id) is None


def test_com_identity_registry_does_not_prune_after_incomplete_poll():
    """A transient partial enumeration must retain unseen tab identities."""
    registry = _ComIdentityRegistry()
    registry.begin_poll()
    first_id = registry.identify(_FakeIdentity("first"), object())
    second_id = registry.identify(_FakeIdentity("second"), object())
    registry.finish_poll(complete=True)

    registry.begin_poll()
    registry.identify(_FakeIdentity("first"), object())
    registry.finish_poll(complete=False)

    assert registry.identity_for(first_id) is not None
    assert registry.identity_for(second_id) is not None
