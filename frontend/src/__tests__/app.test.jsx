import { describe, it, expect, vi, afterEach } from "vitest";
import { render, screen, fireEvent, act, waitFor } from "@testing-library/react";
import { ThemeProvider } from "../theme.jsx";
import VireloApp, { bridgeToState, stateToBridge } from "../app.jsx";

describe("bridgeToState", () => {
  it("Maps Python snake_case keys to React camelCase.", () => {
    const result = bridgeToState({
      enable_snap: true,
      snap_key: "shift",
      restore_key: "ctrl",
      snap_presses: 3,
      snap_interval: 1050,
      width_pct: 76,
      height_pct: 76,
      game_mode_enabled: true,
      ex_auto_size: false,
      run_at_startup: false,
    });
    expect(result.snapEnabled).toBe(true);
    expect(result.snapKey).toBe("SHIFT");
    expect(result.restoreKey).toBe("CTRL");
    expect(result.pressCount).toBe(3);
    expect(result.interval).toBe(1050);
    expect(result.width).toBe(76);
    expect(result.height).toBe(76);
    expect(result.gameMode).toBe(true);
    expect(result.autoSize).toBe(false);
    expect(result.launchLogin).toBe(false);
  });

  it("Uppercases snap_key and restore_key.", () => {
    const result = bridgeToState({ snap_key: "ctrl", restore_key: "alt" });
    expect(result.snapKey).toBe("CTRL");
    expect(result.restoreKey).toBe("ALT");
  });

  it("Applies default values through nullish coalescing.", () => {
    const result = bridgeToState({});
    expect(result.snapEnabled).toBe(true);
    expect(result.snapKey).toBe("SHIFT");
    expect(result.restoreKey).toBe("CTRL");
    expect(result.pressCount).toBe(3);
    expect(result.interval).toBe(1050);
    expect(result.width).toBe(76);
    expect(result.height).toBe(76);
    expect(result.gameMode).toBe(true);
    expect(result.autoSize).toBe(true);
    expect(result.launchLogin).toBe(false);
  });

  it("Handles explicit false values without falling back.", () => {
    const result = bridgeToState({
      enable_snap: false,
      game_mode_enabled: false,
    });
    expect(result.snapEnabled).toBe(false);
    expect(result.gameMode).toBe(false);
  });
});

describe("stateToBridge", () => {
  it("Maps React camelCase keys back to Python snake_case JSON.", () => {
    const json = stateToBridge({
      snapEnabled: true,
      snapKey: "SHIFT",
      restoreKey: "CTRL",
      pressCount: 3,
      interval: 1050,
      width: 76,
      height: 76,
      gameMode: true,
      autoSize: false,
      launchLogin: false,
    });
    const parsed = JSON.parse(json);
    expect(parsed.enable_snap).toBe(true);
    expect(parsed.snap_key).toBe("shift");
    expect(parsed.restore_key).toBe("ctrl");
    expect(parsed.snap_presses).toBe(3);
    expect(parsed.snap_interval).toBe(1050);
    expect(parsed.width_pct).toBe(76);
    expect(parsed.height_pct).toBe(76);
    expect(parsed.game_mode_enabled).toBe(true);
    expect(parsed.ex_auto_size).toBe(false);
    expect(parsed.run_at_startup).toBe(false);
  });

  it("Lowercases key names for the Python bridge.", () => {
    const json = stateToBridge({ snapKey: "SHIFT", restoreKey: "ALT" });
    const parsed = JSON.parse(json);
    expect(parsed.snap_key).toBe("shift");
    expect(parsed.restore_key).toBe("alt");
  });

  it("Returns a valid JSON string.", () => {
    const json = stateToBridge({
      snapEnabled: true,
      snapKey: "SHIFT",
      restoreKey: "CTRL",
      pressCount: 3,
      interval: 1050,
      width: 76,
      height: 76,
      gameMode: true,
      autoSize: false,
      launchLogin: false,
    });
    expect(typeof json).toBe("string");
    expect(() => JSON.parse(json)).not.toThrow();
  });
});

// A minimal bridge mock for full-app renders. Callbacks resolve synchronously
// so tests can assert call ordering without awaiting the event loop.
function makeBridge() {
  const signal = () => ({ connect: vi.fn(), disconnect: vi.fn() });
  return {
    get_settings: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: {} }))),
    settings_changed: signal(),
    dirty_changed: signal(),
    snap_status: signal(),
    capture_status: signal(),
    views_status: signal(),
    views_task_changed: signal(),
    save_settings: vi.fn((json, cb) => cb(JSON.stringify({ ok: true }))),
    commit_draft: vi.fn((cb) => cb(JSON.stringify({ ok: true }))),
    discard_draft: vi.fn((cb) => cb(JSON.stringify({ ok: true }))),
    reset_defaults: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: {} }))),
    test_snap: vi.fn(),
    capture_key: vi.fn(),
    setWindowCommand: vi.fn(),
    apply_details_view: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: { started: true } }))),
    reset_folder_views: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: { started: true } }))),
    restore_folder_views: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: { started: true } }))),
  };
}

function renderApp(bridge) {
  const tweaks = {
    theme: "dark",
    accent: "slate",
    density: "cozy",
    radius: 6,
    sidebarMode: "full",
  };
  return render(
    <ThemeProvider tweaks={tweaks} setTweaks={vi.fn()}>
      <VireloApp bridge={bridge} />
    </ThemeProvider>,
  );
}

describe("Bridge-backed application behavior", () => {
  it("Keeps title-bar controls accessible and outside the draggable spacer.", () => {
    const bridge = makeBridge();
    renderApp(bridge);

    const search = screen.getByRole("button", { name: "Search settings and commands" });
    const minimize = screen.getByRole("button", { name: "Minimize Virelo" });
    const close = screen.getByRole("button", { name: "Close Virelo" });

    expect(search).toHaveStyle({ width: "220px", flex: "0 0 220px" });
    expect(search.previousElementSibling).toHaveStyle({ flex: "1" });
    expect(search.compareDocumentPosition(minimize)).toBe(Node.DOCUMENT_POSITION_FOLLOWING);
    expect(minimize.compareDocumentPosition(close)).toBe(Node.DOCUMENT_POSITION_FOLLOWING);
  });

  it("Allows the main content pane to shrink at the native minimum window width.", () => {
    const bridge = makeBridge();
    const { container } = renderApp(bridge);

    expect(container.querySelector("main")).toHaveStyle({ minWidth: "0" });
  });

  it("Disconnects every application signal subscription on unmount.", () => {
    const bridge = makeBridge();
    const { unmount } = renderApp(bridge);
    const signals = [
      bridge.settings_changed,
      bridge.dirty_changed,
      bridge.snap_status,
      bridge.capture_status,
      bridge.views_status,
      bridge.views_task_changed,
    ];

    unmount();

    for (const signal of signals) {
      const handler = signal.connect.mock.calls[0][0];
      expect(signal.disconnect).toHaveBeenCalledOnce();
      expect(signal.disconnect).toHaveBeenCalledWith(handler);
    }
  });

  it("Keeps folder-view actions busy until the task signal reports completion.", async () => {
    const bridge = makeBridge();
    const { container } = renderApp(bridge);
    fireEvent.click(screen.getByRole("button", { name: "Explorer" }));
    const opener = screen.getByRole("button", { name: "Make Details the default" });
    opener.focus();
    fireEvent.click(opener);
    const appShell = container.querySelector("[data-app-shell]");
    expect(appShell).toHaveAttribute("inert");
    expect(appShell).toHaveAttribute("aria-hidden", "true");
    fireEvent.click(screen.getByRole("button", { name: "Apply Details default" }));

    const working = screen.getByRole("button", { name: "Working..." });
    expect(appShell).not.toHaveAttribute("inert");
    expect(appShell).not.toHaveAttribute("aria-hidden");
    expect(working).toHaveAttribute("aria-disabled", "true");
    await waitFor(() =>
      expect(
        screen.getByText("Folder view task is running. Keep Virelo open until it finishes."),
      ).toHaveFocus(),
    );
    expect(bridge.apply_details_view).toHaveBeenCalledOnce();

    const onViewsStatus = bridge.views_status.connect.mock.calls[0][0];
    act(() => onViewsStatus("Explorer is restarting.", 3000));
    expect(screen.getByRole("button", { name: "Working..." })).toHaveAttribute(
      "aria-disabled",
      "true",
    );

    const onTaskChanged = bridge.views_task_changed.connect.mock.calls[0][0];
    act(() => onTaskChanged(JSON.stringify({ kind: "apply", state: "succeeded" })));
    expect(screen.getByRole("button", { name: "Make Details the default" })).toBeEnabled();
  });

  it("Announces backend messages through the footer live region.", () => {
    const bridge = makeBridge();
    renderApp(bridge);
    const onSnapStatus = bridge.snap_status.connect.mock.calls[0][0];

    act(() => onSnapStatus("Window centered.", 3000));

    expect(screen.getByText("Window centered.")).toHaveAttribute("role", "status");
  });
});

describe("throttled draft writes versus save, discard, and reset", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("Flushes the queued trailing write before committing on save.", () => {
    vi.useFakeTimers();
    const bridge = makeBridge();
    renderApp(bridge);
    // The second switch on the default Window snap page is Game mode.
    const toggle = screen.getAllByRole("switch")[1];
    fireEvent.click(toggle);
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
    // A second change inside the throttle window only queues a trailing write.
    fireEvent.click(toggle);
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
    fireEvent.click(screen.getByText("Save changes"));
    // The queued write goes out before the commit, in that order.
    expect(bridge.save_settings).toHaveBeenCalledTimes(2);
    expect(bridge.commit_draft).toHaveBeenCalledTimes(1);
    expect(bridge.save_settings.mock.invocationCallOrder[1]).toBeLessThan(
      bridge.commit_draft.mock.invocationCallOrder[0],
    );
    // No trailing write fires after the save.
    act(() => {
      vi.advanceTimersByTime(1000);
    });
    expect(bridge.save_settings).toHaveBeenCalledTimes(2);
  });

  it("Cancels the queued trailing write on discard.", () => {
    vi.useFakeTimers();
    const bridge = makeBridge();
    renderApp(bridge);
    const toggle = screen.getAllByRole("switch")[1];
    fireEvent.click(toggle);
    fireEvent.click(toggle);
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
    // Mark the draft dirty so the Discard button appears in the footer.
    const onDirty = bridge.dirty_changed.connect.mock.calls[0][0];
    act(() => onDirty(true));
    fireEvent.click(screen.getByText("Discard"));
    expect(bridge.discard_draft).toHaveBeenCalledTimes(1);
    // The queued write is dropped, so nothing re-dirties the settings.
    act(() => {
      vi.advanceTimersByTime(1000);
    });
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
  });

  it("Cancels the queued trailing write on reset.", () => {
    vi.useFakeTimers();
    const bridge = makeBridge();
    renderApp(bridge);
    const toggle = screen.getAllByRole("switch")[1];
    fireEvent.click(toggle);
    fireEvent.click(toggle);
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
    // Navigate to General and confirm the reset dialog.
    fireEvent.click(screen.getByText("General"));
    fireEvent.click(screen.getByText("Reset"));
    const resetButtons = screen.getAllByText("Reset");
    fireEvent.click(resetButtons[resetButtons.length - 1]);
    expect(bridge.reset_defaults).toHaveBeenCalledTimes(1);
    // The queued write is dropped, so it cannot overwrite the defaults.
    act(() => {
      vi.advanceTimersByTime(1000);
    });
    expect(bridge.save_settings).toHaveBeenCalledTimes(1);
  });
});
