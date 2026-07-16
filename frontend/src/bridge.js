/**
 * QWebChannel bridge client for Virelo.
 *
 * In release mode, QWebChannel is loaded via qrc:///qtwebchannel/qwebchannel.js
 * and the Python VireloBridge QObject is available as channel.objects.bridge.
 *
 * In dev mode (Vite dev server), QWebChannel may not be available. A mock
 * bridge is allowed only in that development build.
 */

let _bridge = null;
let _bridgePromise = null;

function createMockSignal() {
  const handlers = new Set();
  return {
    connect: (handler) => handlers.add(handler),
    disconnect: (handler) => handlers.delete(handler),
    emit: (...args) => handlers.forEach((handler) => handler(...args)),
  };
}

const mockSignals = {
  settings_changed: createMockSignal(),
  theme_applied: createMockSignal(),
  snap_status: createMockSignal(),
  capture_status: createMockSignal(),
  dirty_changed: createMockSignal(),
  views_status: createMockSignal(),
  views_task_changed: createMockSignal(),
};

const MOCK_SETTINGS = {
  snap_key: "shift",
  restore_key: "ctrl",
  enable_snap: true,
  snap_presses: 3,
  snap_interval: 1050,
  width_pct: 76,
  height_pct: 76,
  game_mode_enabled: true,
  ex_auto_size: true,
  run_at_startup: false,
  theme: "dark",
  accent: "slate",
  density: "cozy",
  minimize_to_tray: true,
};

const MOCK_BRIDGE = {
  get_settings: (cb) => cb(JSON.stringify({ ok: true, data: MOCK_SETTINGS })),
  save_settings: (json, cb) => cb(JSON.stringify({ ok: true, applied: JSON.parse(json) })),
  commit_draft: (cb) => cb(JSON.stringify({ ok: true, applied: {} })),
  discard_draft: (cb) => cb(JSON.stringify({ ok: true })),
  has_draft: (cb) => cb(JSON.stringify({ ok: true, data: false })),
  get_snap_enabled: (cb) => cb(JSON.stringify({ ok: true, data: true })),
  test_snap: (cb) => cb(JSON.stringify({ ok: true })),
  capture_key: (target, cb) => cb(JSON.stringify({ ok: true })),
  cancel_capture: (cb) => cb(JSON.stringify({ ok: true })),
  reset_defaults: (cb) => cb(JSON.stringify({ ok: true, data: MOCK_SETTINGS })),
  get_theme_mode: (cb) =>
    cb(JSON.stringify({ ok: true, data: { mode: "dark", effective: "dark" } })),
  get_launch_at_login: (cb) => cb(JSON.stringify({ ok: true, data: false })),
  setWindowCommand: (cmd, cb) => cb(JSON.stringify({ ok: true })),
  apply_details_view: (cb) => runMockViewTask("apply", cb),
  reset_folder_views: (cb) => runMockViewTask("reset", cb),
  restore_folder_views: (cb) => runMockViewTask("restore", cb),
  ...mockSignals,
};

function runMockViewTask(kind, callback) {
  mockSignals.views_task_changed.emit(JSON.stringify({ kind, state: "started" }));
  callback(JSON.stringify({ ok: true, data: { started: true } }));
  setTimeout(() => {
    mockSignals.views_task_changed.emit(JSON.stringify({ kind, state: "succeeded" }));
    mockSignals.views_status.emit("Development preview completed the folder view task.", 3000);
  }, 400);
}

function resolveDevelopmentMock(resolve, reason) {
  console.warn(`[bridge] ${reason} Using the development mock bridge.`);
  _bridge = MOCK_BRIDGE;
  resolve(_bridge);
}

function _initBridge() {
  if (_bridgePromise) return _bridgePromise;

  const pending = new Promise((resolve, reject) => {
    const QWebChannelConstructor = globalThis.QWebChannel;
    const transport = globalThis.qt?.webChannelTransport;
    if (!QWebChannelConstructor) {
      if (import.meta.env.DEV) {
        resolveDevelopmentMock(resolve, "QWebChannel is not available.");
      } else {
        reject(new Error("QWebChannel is not available in this release build."));
      }
      return;
    }

    if (!transport) {
      if (import.meta.env.DEV) {
        resolveDevelopmentMock(resolve, "The Qt WebChannel transport is not available.");
      } else {
        reject(new Error("The Qt WebChannel transport is not available."));
      }
      return;
    }

    try {
      new QWebChannelConstructor(transport, (channel) => {
        _bridge = channel.objects.bridge;
        if (!_bridge) {
          console.error('[bridge] No "bridge" object found in QWebChannel');
          if (import.meta.env.DEV) {
            resolveDevelopmentMock(resolve, "The backend bridge object is missing.");
          } else {
            reject(new Error("The backend bridge object is missing."));
          }
          return;
        }
        resolve(_bridge);
      });
    } catch (error) {
      reject(error);
    }
  });

  _bridgePromise = pending.catch((error) => {
    _bridgePromise = null;
    throw error;
  });

  return _bridgePromise;
}

export function getBridge() {
  return _initBridge();
}
