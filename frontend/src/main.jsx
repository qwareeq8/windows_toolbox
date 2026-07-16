import React from "react";
import { createRoot } from "react-dom/client";
import { ThemeProvider } from "./theme.jsx";
import VireloApp from "./app.jsx";
import { getBridge } from "./bridge.js";

/**
 * Inner component that renders AFTER bridge is ready.
 * This avoids the React hooks violation of calling useState after a conditional return.
 * All hooks in this component are called unconditionally.
 */
function AppWithBridge({ bridge, initialTheme, initialAccent, initialDensity }) {
  const [tweaks, setTweaks] = React.useState({
    theme: initialTheme,
    accent: initialAccent || "slate",
    density: initialDensity || "cozy",
    radius: 6,
    sidebarMode: "full",
  });

  React.useEffect(() => {
    const onThemeApplied = (theme) => {
      setTweaks((prev) => (prev.theme === theme ? prev : { ...prev, theme }));
    };
    const onSettingsChanged = (json) => {
      try {
        const settings = JSON.parse(json);
        setTweaks((prev) => {
          const accent = settings.accent || prev.accent;
          const density = settings.density || prev.density;
          if (accent === prev.accent && density === prev.density) return prev;
          return { ...prev, accent, density };
        });
      } catch (e) {
        console.error("[main] Failed to parse settings_changed for tweaks:", e);
      }
    };

    bridge.theme_applied.connect(onThemeApplied);
    bridge.settings_changed.connect(onSettingsChanged);
    return () => {
      bridge.theme_applied.disconnect?.(onThemeApplied);
      bridge.settings_changed.disconnect?.(onSettingsChanged);
    };
  }, [bridge]);

  const handleSetTweaks = (updates) => {
    setTweaks((prev) => ({ ...prev, ...updates }));
  };

  return (
    <ThemeProvider tweaks={tweaks} setTweaks={handleSetTweaks}>
      <VireloApp bridge={bridge} />
    </ThemeProvider>
  );
}

function BridgeError({ message }) {
  return (
    <div
      role="alert"
      style={{
        width: "100%",
        height: "100%",
        background: "#111113",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "#ECECEE",
        fontFamily: "Arial, Helvetica, sans-serif",
        padding: 32,
      }}
    >
      <div style={{ maxWidth: 520, textAlign: "center" }}>
        <h1 style={{ margin: "0 0 12px", fontSize: 22 }}>Backend connection failed.</h1>
        <p style={{ margin: "0 0 8px", lineHeight: 1.5 }}>
          Virelo cannot safely load or change settings because its Windows backend is unavailable.
        </p>
        <p style={{ margin: "0 0 20px", color: "#B8B8C0", lineHeight: 1.5 }}>{message}</p>
        <button type="button" onClick={() => window.location.reload()}>
          Reload Virelo
        </button>
      </div>
    </div>
  );
}

/**
 * Root component handles bridge initialization only.
 * Uses a single useState + useEffect pair, then conditionally renders
 * either a loading screen or AppWithBridge.
 */
function Root() {
  const [bridgeState, setBridgeState] = React.useState(null);

  React.useEffect(() => {
    // Resolve at most once, from whichever finishes first: the normal
    // bootstrap path, an explicit failure, or the connection timeout.
    let settled = false;
    const finish = (payload) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      setBridgeState(payload);
    };
    const defaults = {
      initialTheme: "dark",
      initialAccent: "slate",
      initialDensity: "cozy",
    };
    // Fail closed if bridge callbacks never fire. A production UI backed by
    // inert mock actions would falsely claim that settings were changed.
    const timer = setTimeout(() => {
      console.warn("[main] Bridge did not respond within 3 seconds.");
      finish({
        error: "The backend did not respond within three seconds. Restart Virelo and try again.",
      });
    }, 3000);
    getBridge()
      .then((b) => {
        b.get_theme_mode((themeResult) => {
          b.get_settings((settingsResult) => {
            try {
              const tr = JSON.parse(themeResult);
              const sr = JSON.parse(settingsResult);
              const themeData = tr.ok && tr.data ? tr.data : { mode: "system", effective: "dark" };
              const settingsData = sr.ok && sr.data ? sr.data : {};
              finish({
                bridge: b,
                initialTheme: themeData.effective || "dark",
                initialAccent: settingsData.accent || "slate",
                initialDensity: settingsData.density || "cozy",
              });
            } catch (e) {
              console.error("[main] Failed to parse initial state:", e);
              finish({ bridge: b, ...defaults });
            }
          });
        });
      })
      .catch((e) => {
        console.error("[main] Bridge initialization failed:", e);
        finish({ error: e instanceof Error ? e.message : String(e) });
      });
    return () => {
      settled = true;
      clearTimeout(timer);
    };
  }, []);

  if (!bridgeState) {
    return (
      <div
        role="status"
        aria-live="polite"
        style={{
          width: "100%",
          height: "100%",
          background: "#111113",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          color: "#ECECEE",
          fontFamily: "Arial, Helvetica, sans-serif",
          fontSize: 14,
        }}
      >
        Loading Virelo...
      </div>
    );
  }

  if (bridgeState.error) return <BridgeError message={bridgeState.error} />;

  return (
    <AppWithBridge
      bridge={bridgeState.bridge}
      initialTheme={bridgeState.initialTheme}
      initialAccent={bridgeState.initialAccent}
      initialDensity={bridgeState.initialDensity}
    />
  );
}

const root = createRoot(document.getElementById("root"));
root.render(<Root />);
