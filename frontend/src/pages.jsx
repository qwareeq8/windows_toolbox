// Pages: Window Snap, Explorer, Shortcuts, General, About.
// Each takes an `app` object with state/setters so Tweaks + palette can
// deep-link into specific settings.

import React from "react";
import { useTokens, useTheme, ACCENTS } from "./theme.jsx";
import {
  Toggle,
  Button,
  Card,
  Row,
  Segmented,
  Stepper,
  Slider,
  Kbd,
  Modal,
} from "./primitives.jsx";
import { Icon } from "./icons.jsx";

// License text shown inline on the About page. Kept in sync with the
// repository LICENSE file.
const MIT_LICENSE = `MIT License

Copyright (c) 2024 Yusuf Qwareeq

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.`;

function ConfirmDialog({
  open,
  onClose,
  title,
  description,
  confirmLabel,
  confirmVariant,
  onConfirm,
}) {
  const t = useTokens();
  const cancelRef = React.useRef(null);
  const titleId = React.useId();
  const descriptionId = React.useId();

  return (
    <Modal
      open={open}
      onClose={onClose}
      labelledBy={titleId}
      describedBy={descriptionId}
      initialFocusRef={cancelRef}
    >
      <div style={{ padding: `${t.cardPad + 4}px ${t.cardPad}px` }}>
        <h2
          id={titleId}
          style={{
            fontSize: 15,
            fontWeight: 600,
            color: t.text,
            margin: "0 0 8px",
          }}
        >
          {title}
        </h2>
        <div
          id={descriptionId}
          style={{
            fontSize: 13,
            color: t.textDim,
            lineHeight: 1.5,
            marginBottom: 20,
          }}
        >
          {description}
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <Button ref={cancelRef} variant="secondary" onClick={onClose}>
            Cancel
          </Button>
          <Button variant={confirmVariant} onClick={onConfirm}>
            {confirmLabel}
          </Button>
        </div>
      </div>
    </Modal>
  );
}

function KeyCapture({ value, target, bridge, label, showStatus }) {
  const t = useTokens();
  const [capturing, setCapturing] = React.useState(false);
  const btnRef = React.useRef(null);

  React.useEffect(() => {
    if (!capturing) return;
    // Tell the backend to release its global keyboard hook. Without this,
    // dismissing the capture UI leaves the hook active and the next key
    // pressed in any application is captured as the new binding.
    const cancelBackend = () => {
      if (bridge.cancel_capture) bridge.cancel_capture(() => {});
    };
    const onStatus = (status) => {
      if (status === "done" || status === "cancelled" || status === "timeout") {
        setCapturing(false);
      }
    };
    bridge.capture_status.connect(onStatus);
    const onKey = (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        cancelBackend();
        setCapturing(false);
      }
    };
    window.addEventListener("keydown", onKey);
    const onMouseDown = (e) => {
      if (btnRef.current && !btnRef.current.contains(e.target)) {
        cancelBackend();
        setCapturing(false);
      }
    };
    const timer = setTimeout(() => {
      document.addEventListener("mousedown", onMouseDown);
    }, 50);
    return () => {
      bridge.capture_status.disconnect(onStatus);
      window.removeEventListener("keydown", onKey);
      document.removeEventListener("mousedown", onMouseDown);
      clearTimeout(timer);
    };
  }, [capturing, bridge]);

  const handleClick = () => {
    if (capturing) return;
    setCapturing(true);
    bridge.capture_key(target, (result) => {
      try {
        const parsed = JSON.parse(result);
        if (!parsed.ok) {
          setCapturing(false);
          showStatus?.(`Key capture failed: ${parsed.error}`, 5000);
        }
      } catch (error) {
        console.error("[capture] Failed to parse capture_key result:", error);
        setCapturing(false);
        showStatus?.("Key capture failed: invalid response from the backend.", 5000);
      }
    });
  };

  return (
    <button
      type="button"
      ref={btnRef}
      onClick={handleClick}
      aria-label={
        capturing
          ? `${label}. Press a key, or press Escape to cancel.`
          : `${label}. Current key: ${value}.`
      }
      aria-pressed={capturing}
      style={{
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        minWidth: 64,
        height: 28,
        padding: "0 12px",
        background: t.isDark ? "rgba(255,255,255,0.06)" : "#fff",
        border: capturing ? `2px solid ${t.accent}` : `1px solid ${t.borderHi}`,
        borderRadius: 4,
        color: capturing ? t.accent : t.text,
        fontSize: 11,
        fontWeight: 600,
        letterSpacing: 0.3,
        fontFamily: t.mono,
        cursor: "pointer",
        boxShadow: t.isDark ? "none" : "0 1px 0 rgba(0,0,0,0.04)",
        transition: "border-color .15s, color .15s",
      }}
    >
      {capturing ? "Press a key..." : value}
    </button>
  );
}

function MonitorPreview({ width, height }) {
  const t = useTokens();
  const mW = 280,
    mH = 170,
    pad = 12;
  const iw = mW - pad * 2,
    ih = mH - pad * 2;
  const wW = (iw * width) / 100,
    wH = (ih * height) / 100;
  const wX = pad + (iw - wW) / 2,
    wY = pad + (ih - wH) / 2;

  // If the UI accent is too dark in light mode (slate), fall back to a sky
  // blue so the rect stays visible against the dark monitor. Otherwise use
  // the accent so the preview reflects the user's choice.
  const darkAccents = ["slate"];
  const { tweaks } = useTheme();
  const rectColor = !t.isDark && darkAccents.includes(tweaks.accent) ? "#8EC4FF" : t.accent;
  const monitorBg = t.isDark ? "#0C0C0E" : "#3B3833";
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "center",
        padding: `${t.cardPad + 6}px ${t.cardPad}px ${t.cardPad + 4}px`,
        background: t.surface,
      }}
    >
      <div style={{ position: "relative" }}>
        <div
          style={{
            position: "absolute",
            top: -22,
            left: 0,
            right: 0,
            textAlign: "center",
            fontSize: 11,
            fontFamily: t.mono,
            color: t.textDim,
            fontWeight: 600,
            letterSpacing: 0.3,
          }}
        >
          {width}% × {height}%
        </div>
        <div
          style={{
            width: mW,
            height: mH,
            borderRadius: 6,
            background: monitorBg,
            border: `1px solid ${t.borderHi}`,
            position: "relative",
            overflow: "hidden",
            boxShadow: "inset 0 1px 0 rgba(255,255,255,0.04)",
          }}
        >
          <svg
            aria-hidden="true"
            width={mW}
            height={mH}
            style={{
              position: "absolute",
              inset: 0,
              opacity: 0.18,
              pointerEvents: "none",
            }}
          >
            <defs>
              <pattern id="mp-grid" width="20" height="20" patternUnits="userSpaceOnUse">
                <path
                  d="M 20 0 L 0 0 0 20"
                  fill="none"
                  stroke="rgba(255,255,255,0.15)"
                  strokeWidth="0.5"
                />
              </pattern>
            </defs>
            <rect width={mW} height={mH} fill="url(#mp-grid)" />
          </svg>

          <div
            style={{
              position: "absolute",
              left: wX,
              top: wY,
              width: wW,
              height: wH,
              background: `${rectColor}33`,
              border: `1.5px solid ${rectColor}`,
              borderRadius: 3,
              transition: "all .2s cubic-bezier(.2,.7,.3,1)",
              boxShadow: `0 0 16px ${rectColor}30`,
            }}
          >
            {wH >= 22 && (
              <div
                style={{
                  height: 10,
                  background: `${rectColor}40`,
                  borderBottom: `1px solid ${rectColor}`,
                  borderRadius: "2px 2px 0 0",
                }}
              />
            )}
          </div>
        </div>
        <div
          style={{
            width: 70,
            height: 4,
            margin: "0 auto",
            background: t.borderHi,
            borderRadius: "0 0 3px 3px",
          }}
        />
        <div
          style={{
            width: 120,
            height: 2,
            margin: "0 auto",
            background: t.border,
            borderRadius: 2,
          }}
        />
      </div>
    </div>
  );
}

function SnapPage({ app }) {
  const t = useTokens();
  return (
    <Pg
      title="Window snap"
      subtitle="Configure snap/restore presets and the keyboard shortcuts that trigger them."
    >
      <Card>
        <Row
          label="Enable snap"
          description="Center the foreground window and resize it when the window supports resizing."
        >
          <Toggle
            label="Enable snap"
            on={app.snapEnabled}
            onChange={(v) => app.set({ snapEnabled: v })}
          />
        </Row>
        <Row label="Game mode" description="Skip snapping while a fullscreen app is in focus." last>
          <Toggle label="Game mode" on={app.gameMode} onChange={(v) => app.set({ gameMode: v })} />
        </Row>
      </Card>

      <Card
        title="Shortcut"
        subtitle="Press the key button to rebind. Tap the bound key repeatedly to trigger."
      >
        <Row label="Snap key">
          <KeyCapture
            value={app.snapKey}
            target="snap"
            bridge={app.bridge}
            label="Capture snap key"
            showStatus={app.showStatus}
          />
        </Row>
        <Row label="Restore key">
          <KeyCapture
            value={app.restoreKey}
            target="restore"
            bridge={app.bridge}
            label="Capture restore key"
            showStatus={app.showStatus}
          />
        </Row>
        <Row label="Press count" description="How many taps trigger the action.">
          <Stepper
            label="Press count"
            value={app.pressCount}
            onChange={(v) => app.set({ pressCount: v })}
            min={1}
            max={10}
          />
        </Row>
        <Row label="Interval" description="Maximum time between taps." last>
          <Stepper
            label="Interval"
            value={app.interval}
            onChange={(v) => app.set({ interval: v })}
            min={100}
            max={5000}
            step={50}
            suffix="ms"
          />
        </Row>
      </Card>

      <Card padding={false}>
        <div
          style={{
            padding: `${t.rowPad + 2}px ${t.cardPad}px`,
            borderBottom: `1px solid ${t.border}`,
            background: t.surface2,
            display: "flex",
            alignItems: "center",
          }}
        >
          <div style={{ flex: 1 }}>
            <div
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: t.text,
                letterSpacing: -0.1,
              }}
            >
              Target size
            </div>
            <div style={{ fontSize: 12, color: t.textDim, marginTop: 2 }}>
              Snapped windows will match this fraction of the current display.
            </div>
          </div>
          <Button variant="ghost" icon={<Icon name="play" size={12} />} onClick={app.onTestSnap}>
            Test snap
          </Button>
        </div>
        <MonitorPreview width={app.width} height={app.height} />
        <div style={{ padding: `4px ${t.cardPad}px` }}>
          <Row label="Width" description={`${app.width}% of screen width`}>
            <Slider
              label="Snap width"
              suffix="%"
              value={app.width}
              onChange={(v) => app.set({ width: v })}
              min={10}
              max={100}
            />
          </Row>
          <Row label="Height" description={`${app.height}% of screen height`} last>
            <Slider
              label="Snap height"
              suffix="%"
              value={app.height}
              onChange={(v) => app.set({ height: v })}
              min={10}
              max={100}
            />
          </Row>
        </div>
      </Card>
    </Pg>
  );
}

function ExplorerPage({ app }) {
  const t = useTokens();
  const [confirmAction, setConfirmAction] = React.useState(null);
  const busy = app.viewsTask ?? null;
  const taskStatusRef = React.useRef(null);

  React.useEffect(() => {
    if (busy === null) return undefined;
    const timer = window.setTimeout(() => taskStatusRef.current?.focus(), 0);
    return () => window.clearTimeout(timer);
  }, [busy]);

  const runViewsAction = (action) => {
    setConfirmAction(null);
    if (busy !== null) {
      return;
    }
    const methods = {
      apply: app.bridge.apply_details_view,
      reset: app.bridge.reset_folder_views,
      restore: app.bridge.restore_folder_views,
    };
    const method = methods[action];
    if (typeof method !== "function") {
      // The backend build in use does not expose the folder view slots yet.
      app.showStatus?.("Folder view changes are not supported by this backend build.", 5000);
      app.setViewsTask?.(null);
      return;
    }
    app.setViewsTask?.(action);
    const onResult = (result) => {
      // The bridge callback only acknowledges that the background task has
      // started. The real success or failure message arrives later through
      // the views_status signal, which app.jsx routes to the footer status.
      try {
        const r = JSON.parse(result);
        if (r.ok) {
          app.showStatus?.(
            action === "apply"
              ? "Writing the Details-view defaults..."
              : action === "reset"
                ? "Resetting the folder-view defaults..."
                : "Restoring the latest folder-view backup...",
            0,
          );
        } else {
          app.setViewsTask?.(null);
          app.showStatus?.(r.error || "Folder view update failed.", 5000);
        }
      } catch (e) {
        console.error("[explorer] Failed to parse folder view result:", e);
        app.setViewsTask?.(null);
        app.showStatus?.("Folder view update failed.", 5000);
      }
    };
    try {
      method(onResult);
    } catch (error) {
      console.error("[explorer] Folder view invocation failed:", error);
      app.setViewsTask?.(null);
      app.showStatus?.("Folder view update could not start.", 5000);
    }
  };

  const confirmCopy = {
    apply: {
      title: "Make Details the default?",
      body: "Windows will clear saved folder layouts. Virelo creates a recovery backup under %LOCALAPPDATA%\\Virelo before changing the registry. Virelo will not close Explorer windows automatically. Restart File Explorer or sign out after the task finishes.",
      confirmLabel: "Apply Details default",
      confirmVariant: "primary",
    },
    reset: {
      title: "Reset folder views?",
      body: "Windows will clear saved folder layouts and restore its defaults. Virelo creates a recovery backup under %LOCALAPPDATA%\\Virelo first. Virelo will not close Explorer windows automatically. Restart File Explorer or sign out after the task finishes.",
      confirmLabel: "Reset folder views",
      confirmVariant: "danger",
    },
    restore: {
      title: "Restore the latest folder view backup?",
      body: "This replaces current folder views with the newest complete Virelo backup under %LOCALAPPDATA%\\Virelo. Virelo will not close Explorer windows automatically. Restart File Explorer or sign out after the task finishes.",
      confirmLabel: "Restore folder views",
      confirmVariant: "primary",
    },
  }[confirmAction] || {
    title: "Folder view action",
    body: "Confirm this folder view action.",
    confirmLabel: "Continue",
    confirmVariant: "primary",
  };

  return (
    <Pg title="Explorer" subtitle="Quality-of-life tweaks for File Explorer's Details view.">
      <Card>
        <Row
          label="Auto-size columns on folder change"
          description="Resize Details view columns to fit each time you navigate."
          last
        >
          <Toggle
            label="Auto-size columns on folder change"
            on={app.autoSize}
            onChange={(v) => app.set({ autoSize: v })}
          />
        </Row>
      </Card>

      <Card title="Default folder view">
        <div style={{ padding: `${t.rowPad}px 0` }}>
          <div
            style={{
              fontSize: 12.5,
              color: t.textDim,
              lineHeight: 1.5,
              maxWidth: 560,
            }}
          >
            Make Details the default view for every folder with a focused workflow inspired by
            WinSetView. Restart File Explorer or sign out after any change.
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 14, flexWrap: "wrap" }}>
            <Button
              variant="primary"
              disabled={busy !== null}
              aria-disabled={busy !== null}
              aria-busy={busy === "apply"}
              onClick={() => setConfirmAction("apply")}
            >
              {busy === "apply" ? "Working..." : "Make Details the default"}
            </Button>
            <Button
              variant="secondary"
              disabled={busy !== null}
              aria-disabled={busy !== null}
              aria-busy={busy === "reset"}
              onClick={() => setConfirmAction("reset")}
            >
              {busy === "reset" ? "Working..." : "Reset folder views to Windows defaults"}
            </Button>
            <Button
              variant="secondary"
              disabled={busy !== null}
              aria-disabled={busy !== null}
              aria-busy={busy === "restore"}
              onClick={() => setConfirmAction("restore")}
            >
              {busy === "restore" ? "Working..." : "Restore latest backup"}
            </Button>
          </div>
          {busy !== null && (
            <div
              ref={taskStatusRef}
              role="status"
              tabIndex={-1}
              style={{ marginTop: 10, color: t.textDim, fontSize: 12.5 }}
            >
              Folder view task is running. Keep Virelo open until it finishes.
            </div>
          )}
        </div>
      </Card>

      <ConfirmDialog
        open={confirmAction !== null}
        onClose={() => setConfirmAction(null)}
        title={confirmCopy.title}
        description={confirmCopy.body}
        confirmLabel={confirmCopy.confirmLabel}
        confirmVariant={confirmCopy.confirmVariant}
        onConfirm={() => runViewsAction(confirmAction)}
      />
    </Pg>
  );
}

function ShortcutsPage({ app }) {
  const t = useTokens();
  const taps = app.pressCount === 1 ? "once" : `${app.pressCount} times`;
  // Restore is a modifier gesture: the restore key is HELD while the snap
  // key is tapped the configured number of times (see snap.py, which checks
  // keyboard.is_pressed(restore_key) when the press target is reached).
  const items = [
    {
      label: "Trigger snap",
      description: `Tap ${app.snapKey} ${taps}.`,
      keys: Array.from({ length: app.pressCount }, () => app.snapKey),
    },
    {
      label: "Restore last snap",
      description: `Hold ${app.restoreKey} while tapping ${app.snapKey} ${taps}.`,
      hold: app.restoreKey,
      keys: Array.from({ length: app.pressCount }, () => app.snapKey),
    },
    { label: "Command palette", keys: ["Ctrl", "K"] },
  ];
  return (
    <Pg title="Shortcuts" subtitle="Global keyboard shortcuts registered by Virelo.">
      <Card padding={false}>
        {items.map((it, i) => (
          <div
            key={i}
            style={{
              display: "flex",
              alignItems: "center",
              padding: `${t.rowPad}px ${t.cardPad}px`,
              borderBottom: i < items.length - 1 ? `1px solid ${t.border}` : "none",
            }}
          >
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 13, color: t.text }}>{it.label}</div>
              {it.description && (
                <div style={{ fontSize: 11.5, color: t.textDim, marginTop: 2 }}>
                  {it.description}
                </div>
              )}
            </div>
            <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
              {it.hold && (
                <>
                  <Kbd>{it.hold}</Kbd>
                  <span
                    style={{
                      color: t.textMuted,
                      fontSize: 10,
                      margin: "0 2px",
                      whiteSpace: "nowrap",
                    }}
                  >
                    (hold)
                  </span>
                  <span
                    style={{
                      color: t.textMuted,
                      fontSize: 11,
                      margin: "0 1px",
                    }}
                  >
                    +
                  </span>
                </>
              )}
              {it.keys.map((k, j) => (
                <React.Fragment key={j}>
                  {j > 0 && (
                    <span
                      style={{
                        color: t.textMuted,
                        fontSize: 11,
                        margin: "0 1px",
                      }}
                    >
                      +
                    </span>
                  )}
                  <Kbd>{k}</Kbd>
                </React.Fragment>
              ))}
            </div>
          </div>
        ))}
      </Card>
    </Pg>
  );
}

function GeneralPage({ app }) {
  const t = useTokens();
  const [confirmReset, setConfirmReset] = React.useState(false);

  return (
    <Pg title="General" subtitle="Application-wide preferences.">
      <Card title="Appearance">
        <Row label="Theme" description="Light or dark surfaces throughout the app.">
          <Segmented
            label="Theme"
            options={[
              { value: "system", label: "System" },
              { value: "light", label: "Light" },
              { value: "dark", label: "Dark" },
            ]}
            value={app.themeMode}
            onChange={(v) => app.set({ themeMode: v })}
          />
        </Row>
        <Row label="Accent color" description="Used for selection, toggles, and primary actions.">
          <div style={{ display: "flex", gap: 6 }}>
            {Object.entries(ACCENTS).map(([k, v]) => (
              <button
                type="button"
                key={k}
                onClick={() => app.set({ accent: k })}
                aria-label={`${k[0].toUpperCase()}${k.slice(1)} accent`}
                aria-pressed={app.accent === k}
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 11,
                  background: t.isDark ? v.dark : v.light,
                  border: app.accent === k ? `2px solid ${t.text}` : `1px solid ${t.border}`,
                  cursor: "pointer",
                  padding: 0,
                }}
              />
            ))}
          </div>
        </Row>
        <Row label="Density" description="Controls spacing throughout the app." last>
          <Segmented
            label="Density"
            options={[
              { value: "compact", label: "Compact" },
              { value: "cozy", label: "Cozy" },
              { value: "comfortable", label: "Comfortable" },
            ]}
            value={app.density}
            onChange={(v) => app.set({ density: v })}
          />
        </Row>
      </Card>

      <Card title="Startup">
        <Row label="Launch at login" description="Start Virelo when you sign in to Windows." last>
          <Toggle
            label="Launch at login"
            on={app.launchLogin}
            onChange={(v) => app.set({ launchLogin: v })}
          />
        </Row>
      </Card>

      <Card title="Advanced">
        <Row
          label="Reset all settings"
          description="Restore every preference to its default value."
          last
        >
          <Button variant="danger" size="sm" onClick={() => setConfirmReset(true)}>
            Reset
          </Button>
        </Row>
      </Card>

      <ConfirmDialog
        open={confirmReset}
        onClose={() => setConfirmReset(false)}
        title="Reset all settings?"
        description="This will restore every preference to its default value. This cannot be undone."
        confirmLabel="Reset"
        confirmVariant="danger"
        onConfirm={() => {
          app.onReset();
          setConfirmReset(false);
        }}
      />
    </Pg>
  );
}

function AboutPage() {
  const t = useTokens();
  const [showLicense, setShowLicense] = React.useState(false);
  return (
    <Pg title="About" subtitle="Virelo -- a tiny utility for snappier windows.">
      <Card padding={false}>
        <div
          style={{
            padding: `${t.cardPad + 4}px ${t.cardPad}px`,
            display: "flex",
            alignItems: "center",
            gap: 16,
          }}
        >
          <div
            style={{
              width: 56,
              height: 56,
              borderRadius: t.radius + 4,
              background: t.accent,
              color: t.accentOn,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 26,
              fontWeight: 700,
              letterSpacing: -1,
            }}
          >
            V
          </div>
          <div style={{ flex: 1 }}>
            <div
              style={{
                fontSize: 18,
                fontWeight: 600,
                color: t.text,
                letterSpacing: -0.3,
              }}
            >
              Virelo
            </div>
            <div style={{ fontSize: 12.5, color: t.textDim, marginTop: 2 }}>
              Version {__APP_VERSION__}
            </div>
          </div>
        </div>
      </Card>
      <Card>
        <Row label="License" description="MIT" last={!showLicense}>
          <Button
            variant="ghost"
            size="sm"
            onClick={() => setShowLicense((v) => !v)}
            aria-expanded={showLicense}
          >
            {showLicense ? "Hide" : "View"}
          </Button>
        </Row>
        {showLicense && (
          <div style={{ padding: `${t.rowPad}px 0` }}>
            <pre
              style={{
                margin: 0,
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
                fontFamily: t.mono,
                fontSize: 11,
                lineHeight: 1.6,
                color: t.textDim,
              }}
            >
              {MIT_LICENSE}
            </pre>
          </div>
        )}
      </Card>
    </Pg>
  );
}

function Pg({ title, subtitle, children }) {
  const t = useTokens();
  return (
    <div>
      <div style={{ marginBottom: t.sectionGap + 6 }}>
        <h1
          style={{
            fontSize: t.titleSize,
            fontWeight: 600,
            letterSpacing: -0.3,
            color: t.text,
          }}
        >
          {title}
        </h1>
        {subtitle && (
          <div
            style={{
              fontSize: 13,
              color: t.textDim,
              marginTop: 4,
              maxWidth: 560,
              lineHeight: 1.5,
            }}
          >
            {subtitle}
          </div>
        )}
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: t.sectionGap }}>{children}</div>
    </div>
  );
}

export { SnapPage, ExplorerPage, ShortcutsPage, GeneralPage, AboutPage, MonitorPreview, Pg };
