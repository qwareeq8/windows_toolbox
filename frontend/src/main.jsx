import React from 'react';
import { createRoot } from 'react-dom/client';
import { ThemeProvider } from './theme.jsx';
import VireloApp from './app.jsx';
import { getBridge } from './bridge.js';

/**
 * Inner component that renders AFTER bridge is ready.
 * This avoids the React hooks violation of calling useState after a conditional return.
 * All hooks in this component are called unconditionally.
 */
function AppWithBridge({ bridge, initialTheme }) {
  const [tweaks, setTweaks] = React.useState({
    theme: initialTheme,
    accent: 'slate',
    density: 'cozy',
    radius: 6,
    sidebarMode: 'full',
  });

  // Subscribe to Python theme changes (system theme polling)
  React.useEffect(() => {
    const handler = (theme) => {
      setTweaks((prev) => ({ ...prev, theme }));
    };
    bridge.theme_applied.connect(handler);
    // QWebChannel signals don't have disconnect in JS — no cleanup needed
  }, [bridge]);

  const handleSetTweaks = (updates) => {
    setTweaks((prev) => {
      const next = { ...prev, ...updates };
      // If theme changed, notify Python
      if (updates.theme && updates.theme !== prev.theme) {
        bridge.apply_theme(updates.theme, () => {});
      }
      return next;
    });
  };

  return (
    <ThemeProvider tweaks={tweaks} setTweaks={handleSetTweaks}>
      <VireloApp bridge={bridge} />
    </ThemeProvider>
  );
}

/**
 * Root component handles bridge initialization only.
 * Uses a single useState + useEffect pair, then conditionally renders
 * either a loading screen or AppWithBridge.
 */
function Root() {
  const [bridgeState, setBridgeState] = React.useState(null);
  // bridgeState is null until ready, then { bridge, initialTheme }

  React.useEffect(() => {
    getBridge().then((b) => {
      b.get_theme_mode((result) => {
        try {
          const r = JSON.parse(result);
          const mode = r.ok && r.data ? r.data : 'dark';
          setBridgeState({
            bridge: b,
            initialTheme: mode === 'light' ? 'light' : 'dark',
          });
        } catch (e) {
          console.error('[main] Failed to parse get_theme_mode result:', e);
          setBridgeState({ bridge: b, initialTheme: 'dark' });
        }
      });
    });
  }, []);

  if (!bridgeState) {
    return (
      <div style={{
        width: '100%', height: '100%', background: '#111113',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        color: '#ECECEE', fontFamily: '"Inter", "Segoe UI", sans-serif', fontSize: 14,
      }}>
        Loading Virelo...
      </div>
    );
  }

  return (
    <AppWithBridge
      bridge={bridgeState.bridge}
      initialTheme={bridgeState.initialTheme}
    />
  );
}

const root = createRoot(document.getElementById('root'));
root.render(<Root />);
