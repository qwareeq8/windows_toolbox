import { describe, it, expect, vi } from "vitest";
import { fireEvent, render, screen } from "@testing-library/react";
import { ThemeProvider } from "../theme.jsx";
import { Badge, Button, Card, Segmented, Slider, Stepper, Toggle } from "../primitives.jsx";

// Wrap the component in ThemeProvider with minimal tweaks.
function renderWithTheme(ui) {
  const tweaks = { theme: "dark", accent: "slate", density: "cozy", radius: 6 };
  return render(
    <ThemeProvider tweaks={tweaks} setTweaks={vi.fn()}>
      {ui}
    </ThemeProvider>,
  );
}

describe("Toggle", () => {
  it("Exposes an accessible name and disabled state.", () => {
    renderWithTheme(<Toggle label="Auto-size columns" on={false} onChange={vi.fn()} disabled />);
    const toggle = screen.getByRole("switch", { name: "Auto-size columns" });
    expect(toggle).toHaveAttribute("aria-checked", "false");
    expect(toggle).toBeDisabled();
  });

  it("Reports the enabled state.", () => {
    renderWithTheme(<Toggle label="Window snap" on onChange={vi.fn()} />);
    expect(screen.getByRole("switch", { name: "Window snap" })).toHaveAttribute(
      "aria-checked",
      "true",
    );
  });
});

describe("Button", () => {
  it("Renders with child text.", () => {
    renderWithTheme(<Button>Click me</Button>);
    expect(screen.getByText("Click me")).toBeInTheDocument();
  });

  it("Renders the primary variant.", () => {
    renderWithTheme(<Button variant="primary">Save</Button>);
    expect(screen.getByText("Save")).toBeInTheDocument();
  });
});

describe("Card", () => {
  it("Renders with a title and children.", () => {
    renderWithTheme(
      <Card title="Test Card">
        <p>Card content</p>
      </Card>,
    );
    expect(screen.getByText("Test Card")).toBeInTheDocument();
    expect(screen.getByText("Card content")).toBeInTheDocument();
  });

  it("Renders without a title.", () => {
    renderWithTheme(
      <Card>
        <p>Just content</p>
      </Card>,
    );
    expect(screen.getByText("Just content")).toBeInTheDocument();
  });
});

describe("Badge", () => {
  it("Renders with the default tone.", () => {
    renderWithTheme(<Badge>NEW</Badge>);
    expect(screen.getByText("NEW")).toBeInTheDocument();
  });

  it("Renders with the accent tone.", () => {
    renderWithTheme(<Badge tone="accent">ON</Badge>);
    expect(screen.getByText("ON")).toBeInTheDocument();
  });
});

describe("Segmented", () => {
  it("Reports the selected option as pressed.", () => {
    renderWithTheme(
      <Segmented
        label="Theme"
        options={["System", "Light", "Dark"]}
        value="Light"
        onChange={vi.fn()}
      />,
    );
    expect(screen.getByRole("group", { name: "Theme" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Light" })).toHaveAttribute("aria-pressed", "true");
  });
});

describe("Stepper", () => {
  it("Names its controls and enforces its bounds.", () => {
    renderWithTheme(<Stepper label="Press count" value={1} min={1} max={3} onChange={vi.fn()} />);
    expect(screen.getByRole("button", { name: "Decrease Press count" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Increase Press count" })).toBeEnabled();
    expect(screen.getByRole("status", { name: "Press count: 1" })).toBeInTheDocument();
  });
});

describe("Slider", () => {
  it("Supports standard keyboard adjustments and exposes its value.", () => {
    const onChange = vi.fn();
    renderWithTheme(
      <Slider label="Snap width" value={50} min={10} max={100} suffix="%" onChange={onChange} />,
    );
    const slider = screen.getByRole("slider", { name: "Snap width" });
    expect(slider).toHaveAttribute("aria-valuetext", "50%");

    fireEvent.keyDown(slider, { key: "ArrowUp" });
    expect(onChange).toHaveBeenLastCalledWith(51);
    fireEvent.keyDown(slider, { key: "Home" });
    expect(onChange).toHaveBeenLastCalledWith(10);
    fireEvent.keyDown(slider, { key: "End" });
    expect(onChange).toHaveBeenLastCalledWith(100);
  });
});
