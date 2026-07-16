// Command palette (Ctrl/Cmd+K) for Virelo.

import React from "react";
import { useTokens } from "./theme.jsx";
import { Kbd, Modal } from "./primitives.jsx";
import { Icon } from "./icons.jsx";

function CommandPalette({ open, onClose, app, setNav, onTestSnap, onSave }) {
  const t = useTokens();
  const [q, setQ] = React.useState("");
  const [idx, setIdx] = React.useState(0);
  const inputRef = React.useRef(null);
  const listboxId = React.useId();

  React.useEffect(() => {
    if (open) {
      setQ("");
      setIdx(0);
    }
  }, [open]);

  const commands = React.useMemo(
    () => [
      {
        grp: "Navigate",
        label: "Go to Window snap",
        run: () => setNav("snap"),
        icon: "snap",
      },
      {
        grp: "Navigate",
        label: "Go to Explorer",
        run: () => setNav("exp"),
        icon: "folder",
      },
      {
        grp: "Navigate",
        label: "Go to Shortcuts",
        run: () => setNav("keys"),
        icon: "keyb",
      },
      {
        grp: "Navigate",
        label: "Go to General",
        run: () => setNav("gen"),
        icon: "general",
      },
      {
        grp: "Navigate",
        label: "Go to About",
        run: () => setNav("about"),
        icon: "about",
      },
      {
        grp: "Actions",
        label: "Test snap",
        run: () => onTestSnap?.(),
        icon: "play",
        kbd: "⏎",
      },
      {
        grp: "Actions",
        label: "Save changes",
        run: () => onSave?.(),
        icon: "check",
      },
      {
        grp: "Actions",
        label: app.snapEnabled ? "Disable snap" : "Enable snap",
        run: () => app.set({ snapEnabled: !app.snapEnabled }),
        icon: "dot",
      },
      {
        grp: "Actions",
        label: app.gameMode ? "Disable game mode" : "Enable game mode",
        run: () => app.set({ gameMode: !app.gameMode }),
        icon: "dot",
      },
      {
        grp: "Theme",
        label: "Theme: System",
        run: () => app.set({ themeMode: "system" }),
        icon: "spark",
      },
      {
        grp: "Theme",
        label: "Theme: Light",
        run: () => app.set({ themeMode: "light" }),
        icon: "spark",
      },
      {
        grp: "Theme",
        label: "Theme: Dark",
        run: () => app.set({ themeMode: "dark" }),
        icon: "spark",
      },
      {
        grp: "Theme",
        label: "Accent: Slate",
        run: () => app.set({ accent: "slate" }),
        icon: "dot",
      },
      {
        grp: "Theme",
        label: "Accent: Teal",
        run: () => app.set({ accent: "teal" }),
        icon: "dot",
      },
      {
        grp: "Theme",
        label: "Accent: Blue",
        run: () => app.set({ accent: "blue" }),
        icon: "dot",
      },
      {
        grp: "Theme",
        label: "Accent: Rust",
        run: () => app.set({ accent: "rust" }),
        icon: "dot",
      },
      {
        grp: "Theme",
        label: "Accent: Purple",
        run: () => app.set({ accent: "purple" }),
        icon: "dot",
      },
    ],
    [app, setNav, onTestSnap, onSave],
  );

  const filtered = q.trim()
    ? commands.filter((c) => c.label.toLowerCase().includes(q.toLowerCase()))
    : commands;

  const groups = {};
  filtered.forEach((c) => {
    (groups[c.grp] = groups[c.grp] || []).push(c);
  });

  React.useEffect(() => {
    setIdx(0);
  }, [q]);
  const onInputKeyDown = (event) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      setIdx((current) => (filtered.length === 0 ? 0 : Math.min(filtered.length - 1, current + 1)));
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      setIdx((current) => Math.max(0, current - 1));
    }
    if (event.key === "Enter" && filtered[idx]) {
      event.preventDefault();
      filtered[idx].run();
      onClose();
    }
  };

  if (!open) return null;
  return (
    <Modal
      open={open}
      onClose={onClose}
      ariaLabel="Command palette"
      initialFocusRef={inputRef}
      width={480}
      align="start"
    >
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          maxHeight: 420,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            padding: "12px 14px",
            borderBottom: `1px solid ${t.border}`,
          }}
        >
          <span style={{ color: t.textDim }}>
            <Icon name="search" size={15} />
          </span>
          <input
            ref={inputRef}
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onKeyDown={onInputKeyDown}
            placeholder="Search settings, jump to..."
            aria-label="Search commands"
            role="combobox"
            aria-expanded="true"
            aria-controls={listboxId}
            aria-autocomplete="list"
            aria-activedescendant={filtered[idx] ? `${listboxId}-option-${idx}` : undefined}
            style={{
              flex: 1,
              border: "none",
              outline: "none",
              background: "transparent",
              color: t.text,
              fontSize: 14,
              fontFamily: "inherit",
            }}
          />
          <Kbd>Esc</Kbd>
        </div>
        <div
          id={listboxId}
          role="listbox"
          aria-label="Available commands"
          style={{ flex: 1, overflowY: "auto", padding: "6px 0" }}
        >
          {filtered.length === 0 && (
            <div
              role="status"
              style={{
                padding: 24,
                textAlign: "center",
                color: t.textMuted,
                fontSize: 13,
              }}
            >
              No results for "{q}"
            </div>
          )}
          {Object.entries(groups).map(([grp, items]) => (
            <div key={grp} role="group" aria-label={grp}>
              <div
                style={{
                  padding: "6px 14px 2px",
                  fontSize: 10,
                  fontWeight: 600,
                  color: t.textMuted,
                  textTransform: "uppercase",
                  letterSpacing: 0.8,
                }}
              >
                {grp}
              </div>
              {items.map((c) => {
                // Capture a per-item flat index so each row's onMouseEnter
                // closure highlights that row instead of sharing one mutable
                // counter that ends at the last item.
                const itemIdx = filtered.indexOf(c);
                const active = itemIdx === idx;
                return (
                  <div
                    key={c.label}
                    id={`${listboxId}-option-${itemIdx}`}
                    role="option"
                    aria-selected={active}
                    onClick={() => {
                      c.run();
                      onClose();
                    }}
                    onMouseEnter={() => setIdx(itemIdx)}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 10,
                      padding: "8px 14px",
                      cursor: "pointer",
                      background: active ? t.accentBg : "transparent",
                      color: active ? t.accent : t.text,
                    }}
                  >
                    <span style={{ color: active ? t.accent : t.textDim }}>
                      <Icon name={c.icon} size={13} />
                    </span>
                    <span style={{ flex: 1, fontSize: 13 }}>{c.label}</span>
                    {c.kbd && <Kbd>{c.kbd}</Kbd>}
                  </div>
                );
              })}
            </div>
          ))}
        </div>
        <div
          style={{
            padding: "8px 14px",
            borderTop: `1px solid ${t.border}`,
            background: t.surface2,
            display: "flex",
            alignItems: "center",
            gap: 12,
            fontSize: 11,
            color: t.textMuted,
          }}
        >
          <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <Kbd>{"↑"}</Kbd>
            <Kbd>{"↓"}</Kbd> navigate
          </span>
          <span style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <Kbd>{"⏎"}</Kbd> select
          </span>
          <div style={{ flex: 1 }} />
          <span>Virelo {__APP_VERSION__}</span>
        </div>
      </div>
    </Modal>
  );
}

export { CommandPalette };
