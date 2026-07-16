import { describe, it, expect, vi } from "vitest";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ThemeProvider } from "../theme.jsx";
import { ShortcutsPage, ExplorerPage, SnapPage } from "../pages.jsx";

// Wrap the component under test in a ThemeProvider with minimal tweaks.
function renderWithTheme(ui) {
  const tweaks = { theme: "dark", accent: "slate", density: "cozy", radius: 6 };
  return render(
    <ThemeProvider tweaks={tweaks} setTweaks={vi.fn()}>
      {ui}
    </ThemeProvider>,
  );
}

function makeApp(overrides = {}) {
  return {
    snapEnabled: true,
    snapKey: "SHIFT",
    restoreKey: "CTRL",
    pressCount: 3,
    interval: 1050,
    width: 76,
    height: 76,
    gameMode: true,
    autoSize: true,
    launchLogin: false,
    accent: "slate",
    density: "cozy",
    themeMode: "system",
    set: vi.fn(),
    onTestSnap: vi.fn(),
    onReset: vi.fn(),
    showStatus: vi.fn(),
    viewsTask: null,
    setViewsTask: vi.fn(),
    bridge: {
      capture_key: vi.fn(),
      capture_status: { connect: vi.fn(), disconnect: vi.fn() },
      apply_details_view: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: { started: true } }))),
      reset_folder_views: vi.fn((cb) => cb(JSON.stringify({ ok: true, data: { started: true } }))),
      restore_folder_views: vi.fn((cb) =>
        cb(JSON.stringify({ ok: true, data: { started: true } })),
      ),
    },
    ...overrides,
  };
}

describe("ShortcutsPage", () => {
  it("Renders one snap-key chip per press in both snap and restore rows.", () => {
    renderWithTheme(<ShortcutsPage app={makeApp({ pressCount: 5 })} />);
    // The trigger row and the restore row each show the snap key once per
    // press, so the snap key appears twice per configured press.
    expect(screen.getAllByText("SHIFT")).toHaveLength(10);
  });

  it("Shows the restore key as a single held modifier instead of repeated presses.", () => {
    renderWithTheme(<ShortcutsPage app={makeApp({ pressCount: 5 })} />);
    expect(screen.getAllByText("CTRL")).toHaveLength(1);
    expect(screen.getByText("(hold)")).toBeInTheDocument();
    expect(screen.getByText("Hold CTRL while tapping SHIFT 5 times.")).toBeInTheDocument();
  });

  it("Uses singular copy and a single chip pair when pressCount is 1.", () => {
    renderWithTheme(<ShortcutsPage app={makeApp({ pressCount: 1 })} />);
    expect(screen.getAllByText("SHIFT")).toHaveLength(2);
    expect(screen.getAllByText("CTRL")).toHaveLength(1);
    expect(screen.getByText("Hold CTRL while tapping SHIFT once.")).toBeInTheDocument();
  });
});

describe("SnapPage size sliders", () => {
  it("Limits the width and height sliders to the backend range 10 to 100.", () => {
    renderWithTheme(<SnapPage app={makeApp()} />);
    const sliders = screen.getAllByRole("slider");
    expect(sliders).toHaveLength(2);
    for (const slider of sliders) {
      expect(slider).toHaveAttribute("aria-valuemin", "10");
      expect(slider).toHaveAttribute("aria-valuemax", "100");
    }
  });
});

describe("ExplorerPage default folder view", () => {
  it("Renders the section title, body copy, and all recovery actions.", () => {
    renderWithTheme(<ExplorerPage app={makeApp()} />);
    expect(screen.getByText("Default folder view")).toBeInTheDocument();
    expect(screen.getByText(/inspired by\s+WinSetView/)).toBeInTheDocument();
    expect(screen.getByText("Make Details the default")).toBeInTheDocument();
    expect(screen.getByText("Reset folder views to Windows defaults")).toBeInTheDocument();
    expect(screen.getByText("Restore latest backup")).toBeInTheDocument();
  });

  it("Shows an in-progress message instead of success when apply starts.", async () => {
    const app = makeApp();
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Make Details the default"));
    expect(screen.getByText("Make Details the default?")).toBeInTheDocument();
    expect(screen.getByText(/will not close Explorer windows/i)).toBeInTheDocument();
    expect(screen.getByText(/clear saved folder layouts/i)).toBeInTheDocument();
    expect(screen.getByText(/%LOCALAPPDATA%\\Virelo/)).toBeInTheDocument();
    await waitFor(() => expect(screen.getByText("Cancel")).toHaveFocus());
    await user.click(screen.getByRole("button", { name: "Apply Details default" }));
    expect(app.bridge.apply_details_view).toHaveBeenCalledTimes(1);
    // The bridge callback only acknowledges a start. The completion message
    // arrives later through the views_status signal, so no success text may
    // be announced here.
    expect(app.showStatus).toHaveBeenCalledWith("Writing the Details-view defaults...", 0);
    expect(app.showStatus).not.toHaveBeenCalledWith("Details view applied.", expect.anything());
    expect(app.setViewsTask).toHaveBeenCalledWith("apply");
  });

  it("Shows an in-progress message instead of success when reset starts.", async () => {
    const app = makeApp();
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Reset folder views to Windows defaults"));
    expect(screen.getByText("Reset folder views?")).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Reset folder views" }));
    expect(app.bridge.reset_folder_views).toHaveBeenCalledTimes(1);
    expect(app.showStatus).toHaveBeenCalledWith("Resetting the folder-view defaults...", 0);
    expect(app.showStatus).not.toHaveBeenCalledWith("Folder views reset.", expect.anything());
    expect(app.setViewsTask).toHaveBeenCalledWith("reset");
  });

  it("Restores the latest verified backup through a confirmed action.", async () => {
    const app = makeApp();
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Restore latest backup"));
    expect(
      screen.getByRole("dialog", { name: "Restore the latest folder view backup?" }),
    ).toBeInTheDocument();
    await user.click(screen.getByRole("button", { name: "Restore folder views" }));
    expect(app.bridge.restore_folder_views).toHaveBeenCalledTimes(1);
    expect(app.setViewsTask).toHaveBeenCalledWith("restore");
    expect(app.showStatus).toHaveBeenCalledWith("Restoring the latest folder-view backup...", 0);
  });

  it("Disables every folder-view action while a task is running.", async () => {
    const app = makeApp({ viewsTask: "apply" });
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    const activeButton = screen.getByText("Working...").closest("button");
    expect(activeButton).toBeDisabled();
    expect(
      screen.getByText("Reset folder views to Windows defaults").closest("button"),
    ).toBeDisabled();
    expect(screen.getByText("Restore latest backup").closest("button")).toBeDisabled();
    await user.click(activeButton);
    expect(screen.queryByRole("dialog")).not.toBeInTheDocument();
    expect(app.bridge.apply_details_view).not.toHaveBeenCalled();
  });

  it("Does not call the bridge when the dialog is cancelled.", async () => {
    const app = makeApp();
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Make Details the default"));
    await user.click(screen.getByText("Cancel"));
    expect(app.bridge.apply_details_view).not.toHaveBeenCalled();
    expect(screen.queryByText("Make Details the default?")).not.toBeInTheDocument();
  });

  it("Traps dialog focus, closes on Escape, and restores focus to the opener.", async () => {
    const app = makeApp();
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    const opener = screen.getByRole("button", {
      name: "Make Details the default",
    });

    await user.click(opener);
    const dialog = screen.getByRole("dialog", {
      name: "Make Details the default?",
    });
    const cancel = screen.getByRole("button", { name: "Cancel" });
    const confirm = screen.getByRole("button", {
      name: "Apply Details default",
    });
    await waitFor(() => expect(cancel).toHaveFocus());

    await user.tab({ shift: true });
    expect(confirm).toHaveFocus();
    await user.tab();
    expect(cancel).toHaveFocus();
    await user.keyboard("{Escape}");

    expect(dialog).not.toBeInTheDocument();
    expect(opener).toHaveFocus();
  });

  it("Surfaces a backend error through showStatus.", async () => {
    const app = makeApp();
    app.bridge.apply_details_view = vi.fn((cb) =>
      cb(JSON.stringify({ ok: false, error: "Explorer restart failed." })),
    );
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Make Details the default"));
    await user.click(screen.getByRole("button", { name: "Apply Details default" }));
    expect(app.showStatus).toHaveBeenCalledWith("Explorer restart failed.", 5000);
    expect(app.setViewsTask).toHaveBeenLastCalledWith(null);
  });

  it("Reports an unsupported backend when the method is missing.", async () => {
    const app = makeApp();
    delete app.bridge.apply_details_view;
    renderWithTheme(<ExplorerPage app={app} />);
    const user = userEvent.setup();
    await user.click(screen.getByText("Make Details the default"));
    await user.click(screen.getByRole("button", { name: "Apply Details default" }));
    expect(app.showStatus).toHaveBeenCalledWith(
      "Folder view changes are not supported by this backend build.",
      5000,
    );
  });
});
