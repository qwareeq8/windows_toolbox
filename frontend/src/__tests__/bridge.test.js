import { afterEach, describe, expect, it, vi } from "vitest";

describe("QWebChannel bridge initialization", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllEnvs();
    vi.unstubAllGlobals();
    vi.resetModules();
  });

  it("Rejects a missing backend bridge in a production build.", async () => {
    vi.stubEnv("DEV", false);
    vi.stubGlobal("QWebChannel", undefined);
    vi.stubGlobal("qt", undefined);
    const { getBridge } = await import("../bridge.js");

    await expect(getBridge()).rejects.toThrow(
      "QWebChannel is not available in this release build.",
    );
  });

  it("Emits task start and completion states from the development mock.", async () => {
    vi.useFakeTimers();
    vi.stubEnv("DEV", true);
    vi.stubGlobal("QWebChannel", undefined);
    vi.stubGlobal("qt", undefined);
    const { getBridge } = await import("../bridge.js");
    const bridge = await getBridge();
    const states = [];
    const onTaskChanged = (json) => states.push(JSON.parse(json));
    bridge.views_task_changed.connect(onTaskChanged);

    const acknowledgement = vi.fn();
    bridge.apply_details_view(acknowledgement);
    expect(states).toEqual([{ kind: "apply", state: "started" }]);
    expect(JSON.parse(acknowledgement.mock.calls[0][0])).toMatchObject({
      ok: true,
      data: { started: true },
    });

    vi.advanceTimersByTime(400);
    expect(states).toEqual([
      { kind: "apply", state: "started" },
      { kind: "apply", state: "succeeded" },
    ]);
    bridge.views_task_changed.disconnect(onTaskChanged);
  });
});
