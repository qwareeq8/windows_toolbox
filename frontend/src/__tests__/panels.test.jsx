import { describe, it, expect, vi } from "vitest";
import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ThemeProvider } from "../theme.jsx";
import { CommandPalette } from "../panels.jsx";

// Wrap the component in ThemeProvider with minimal tweaks.
function renderPalette(props = {}) {
  const tweaks = { theme: "dark", accent: "slate", density: "cozy", radius: 6 };
  const setTweaks = vi.fn();
  const app = {
    snapEnabled: true,
    gameMode: true,
    set: vi.fn(),
    ...props.app,
  };
  const defaults = {
    open: true,
    onClose: vi.fn(),
    app,
    setNav: vi.fn(),
    onTestSnap: vi.fn(),
    onSave: vi.fn(),
    onReset: vi.fn(),
    ...props,
  };
  return render(
    <ThemeProvider tweaks={tweaks} setTweaks={setTweaks}>
      <CommandPalette {...defaults} />
    </ThemeProvider>,
  );
}

describe("CommandPalette", () => {
  it("Renders the search input when open.", async () => {
    renderPalette();
    const input = screen.getByRole("combobox", { name: "Search commands" });
    expect(input).toBeInTheDocument();
    await waitFor(() => expect(input).toHaveFocus());
    expect(screen.getByRole("dialog", { name: "Command palette" })).toBeInTheDocument();
    expect(screen.getByRole("listbox", { name: "Available commands" })).toBeInTheDocument();
    expect(screen.getAllByRole("option").length).toBeGreaterThan(0);
  });

  it("Shows all commands when no filter is applied.", () => {
    renderPalette();
    // Navigate group commands should appear
    expect(screen.getByText("Go to Window snap")).toBeInTheDocument();
    expect(screen.getByText("Go to Explorer")).toBeInTheDocument();
    // Action group commands should appear
    expect(screen.getByText("Test snap")).toBeInTheDocument();
    expect(screen.getByText("Save changes")).toBeInTheDocument();
  });

  it("Filters commands when typing a search query.", async () => {
    renderPalette();
    const user = userEvent.setup();
    const input = screen.getByPlaceholderText(/search/i);
    await user.type(input, "snap");
    // "Test snap" and "Go to Window snap" should be visible
    expect(screen.getByText("Test snap")).toBeInTheDocument();
    expect(screen.getByText("Go to Window snap")).toBeInTheDocument();
    // Unrelated commands should be filtered out
    expect(screen.queryByText("Go to Explorer")).not.toBeInTheDocument();
    expect(screen.queryByText("Go to About")).not.toBeInTheDocument();
  });

  it("Shows a no-results message for an unmatched filter.", async () => {
    renderPalette();
    const user = userEvent.setup();
    const input = screen.getByPlaceholderText(/search/i);
    await user.type(input, "zzzznonexistent");
    expect(screen.getByText(/no results/i)).toBeInTheDocument();
  });

  it("Does not render when open is false.", () => {
    renderPalette({ open: false });
    expect(screen.queryByPlaceholderText(/search/i)).not.toBeInTheDocument();
  });

  it("Highlights the hovered item instead of the last item.", async () => {
    const setNav = vi.fn();
    renderPalette({ setNav });
    const user = userEvent.setup();
    const first = screen.getByText("Go to Window snap").closest("div");
    const second = screen.getByText("Go to Explorer").closest("div");
    await user.hover(second);
    // The hovered row becomes active and the previously active row clears.
    expect(second.style.background).not.toBe("transparent");
    expect(first.style.background).toBe("transparent");
    // Enter runs the hovered command, proving the hover index is correct.
    await user.keyboard("{Enter}");
    expect(setNav).toHaveBeenCalledWith("exp");
  });
});
