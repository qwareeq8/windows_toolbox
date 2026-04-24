# Domain Pitfalls

**Domain:** PySide6/QWebEngine/PyInstaller hybrid desktop app -- cleanup and hardening
**Researched:** 2026-04-24

## Critical Pitfalls

Mistakes that cause crashes, broken builds, or security holes during cleanup/hardening of this kind of app.

---

### Pitfall 1: PyInstaller Builds Without Frontend Assets

**What goes wrong:** The PyInstaller spec file references `("frontend/dist", "frontend/dist")` as bundled data, but the build script (`build-installer.ps1`) runs `PyInstaller` directly without first running `npm ci && npm run build` in the frontend directory. If `frontend/dist/` does not exist or is stale, PyInstaller either fails with a missing-path error or bundles an outdated UI. The frozen app then shows a blank white page or loads a wrong version of the frontend.

**Why it happens:** The Python and JS build steps evolved independently. The spec file was written assuming `frontend/dist/` already exists, and nobody added the npm step to the build script.

**Consequences:** Silent production bugs (stale UI) or build failures that are confusing to diagnose. The existing `webview.py:68-72` fallback path tries a second location but neither will work if the build was never run.

**Prevention:**
1. The build script must run `npm ci && npm run build` inside `frontend/` before invoking PyInstaller. Fail the build if `frontend/dist/index.html` does not exist after that step.
2. Add a pre-build assertion in the spec file or a wrapper script: `assert os.path.isfile("frontend/dist/index.html")`.
3. In CI, make frontend build a prerequisite job/step that runs before the PyInstaller step.

**Detection:** Build the app from a clean checkout. If the app starts with a blank page or a "file not found" log entry, this pitfall was hit.

**Phase relevance:** Build pipeline phase -- must be fixed as the very first infrastructure change.

---

### Pitfall 2: QWebEngine Navigation Not Restricted

**What goes wrong:** The `VireloWebView` at `webview.py:109` sets `LocalContentCanAccessRemoteUrls = True` and does not override `acceptNavigationRequest`. If a future React bug or dependency injects an anchor tag, script redirect, or iframe pointing to an external URL, the embedded browser silently navigates to the internet. In an admin-elevated process with a global keyboard hook, this is a privilege escalation vector -- a remote page could interact with the QWebChannel bridge.

**Why it happens:** During rapid prototyping, `LocalContentCanAccessRemoteUrls = True` was enabled to make Vite HMR work in dev mode, and no production-mode lockdown was added.

**Consequences:** An attacker who can influence any URL loaded in the WebEngine (via a compromised npm dependency, XSS in a future feature, or a crafted file:// URL) gains access to every `@Slot` method on the bridge -- which includes toggling startup shortcuts, changing settings, and triggering key capture sessions. All of this runs as administrator.

**Prevention:**
1. Override `acceptNavigationRequest` in `VireloWebPage` to reject all navigations except the initial `file://` load (or `localhost:5173` in dev mode). Return `False` for everything else.
2. Set `LocalContentCanAccessRemoteUrls = False` in production mode. Only enable it when `_is_dev_mode()` returns `True`.
3. Consider registering a custom URL scheme (`virelo://`) with `QWebEngineUrlScheme` and the `LocalScheme` + `Secure` flags instead of using raw `file://`. This blocks cross-origin access by design.

**Detection:** In the running app, open the browser dev tools (if enabled) and try `window.location = "https://example.com"`. If it navigates, the pitfall is present.

**Phase relevance:** WebEngine hardening phase. Must happen before any public release or GitHub publication.

---

### Pitfall 3: QWebChannel Bridge Exposes All Public Methods of QObject

**What goes wrong:** When a QObject is registered with `QWebChannel.registerObject("bridge", bridge)`, Qt publishes all public methods, properties, and signals of that object to the JavaScript side -- not just the `@Slot`-decorated ones. If the bridge class inherits or defines any public method (including inherited QObject methods like `deleteLater`, `setProperty`, `findChild`), those become callable from JS. The `VireloBridge` currently calls `self._main_window._start_key_capture()` and other private MainWindow methods, meaning if the bridge ever accidentally exposed `_main_window`, JS could call arbitrary private methods on the main window.

**Why it happens:** QWebChannel's publishing is generous by design -- it mirrors the C++ QObject introspection system. Developers assume only `@Slot`-decorated methods are visible, but that is not how Qt works.

**Consequences:** Unintended method exposure. Currently mitigated because `_main_window` is a Python attribute (not a Qt property), but any refactoring that turns it into a Q_PROPERTY or public method creates an attack surface.

**Prevention:**
1. Keep the bridge class minimal -- only `@Slot`-decorated methods with explicit type signatures.
2. Never store the MainWindow reference as a Qt property on the bridge. The current `self._main_window = None` pattern (plain Python attribute) is correct -- do not "improve" it by making it a property.
3. Audit the bridge after every refactoring pass to confirm no new public methods leaked.
4. Consider `blockUpdates` on the channel during sensitive operations.

**Detection:** In the browser console, inspect `channel.objects.bridge` and enumerate all available methods. Anything beyond the explicitly defined slots is a leak.

**Phase relevance:** Bridge refactoring phase and WebEngine hardening phase.

---

### Pitfall 4: Splitting main.py Breaks Signal/Slot Wiring and Object Lifetimes

**What goes wrong:** Extracting classes from a 1451-line monolithic `main.py` into separate modules introduces three failure modes:
- **Circular imports:** `main.py` imports `bridge.py`, which already calls `self._main_window._*` private methods. If the extracted modules import each other, Python raises `ImportError` at startup.
- **Signal disconnection:** Signals connected via `app.aboutToQuit.connect(lambda: ...)` at line 1440-1442 capture `win` by closure. If `win` is garbage-collected because the reference was lost during refactoring, the lambda fires on a deleted C++ object: `RuntimeError: Internal C++ object already deleted`.
- **Qt parent chain breaks:** Widgets that relied on implicit parent relationships (via being created in the same file scope as their parent) may lose their parents when moved to a new module, causing them to become top-level windows or be garbage-collected prematurely.

**Why it happens:** Python's GC and Qt's C++ ownership model are fundamentally different. In a monolithic file, everything shares the same scope and lifetime. Splitting introduces module-level scope boundaries that can break implicit lifetime guarantees.

**Consequences:** Crashes at shutdown (the most common symptom), widgets disappearing, or signals silently disconnecting and features stopping to work without any error message.

**Prevention:**
1. Extract pure logic first (win32 helpers, path canonicalization, startup shortcut functions). These have no Qt dependencies and cannot cause signal/slot issues.
2. For Qt classes (ShiftSnapRestore, MainWindow), keep them in files that import from the extracted utility modules, not the other way around.
3. Use explicit `parent=` arguments on every QObject construction. Never rely on "it works because it's in the same file."
4. After each extraction, run the app and verify: startup, settings change, key capture, theme toggle, shutdown. These exercise the most signal-heavy codepaths.
5. Keep `MainWindow` in `main.py` initially. It has the most signal connections and is the hardest to extract safely.

**Detection:** `RuntimeError: Internal C++ object already deleted` in logs. App crashes on close. Features silently stop working after settings changes.

**Phase relevance:** Module splitting phase. This is the highest-risk refactoring task.

---

### Pitfall 5: COM Apartment Threading Violations During Refactoring

**What goes wrong:** The Explorer autosize worker initializes COM as STA (`COINIT_APARTMENTTHREADED`) at `workers.py:731` and caches a `Shell.Application` COM object. If refactoring moves any COM call to a different thread (e.g., the main Qt thread, or a new helper thread), the call crosses apartment boundaries without marshaling, causing:
- `pywintypes.com_error: (-2147417842, 'The application called an interface that was marshalled for a different thread.')`
- Hard crashes with no Python traceback (the error originates in the COM runtime).

**Why it happens:** COM apartment rules are invisible in Python code. Nothing in the syntax prevents calling a COM object from the wrong thread. The `workers.py` code works only because the entire COM lifecycle (init, use, uninit) happens on a single `QThread`. Refactoring that breaks this co-location breaks COM.

**Consequences:** Hard crashes, data corruption in Explorer column autosize, or silent failures where COM calls return empty results. These are extremely difficult to reproduce because they depend on thread scheduling.

**Prevention:**
1. Document the COM threading contract at the top of any module that uses COM: "All COM objects in this module must be created and used on the same STA thread."
2. When extracting `ExplorerAutosizeEngine` into its own module, keep the COM initialization and the Shell.Application cache in the same class/method that uses them. Do not pass COM objects across thread boundaries.
3. Never import `pythoncom` or call `CoInitialize` in a module's top-level scope -- it must be called inside the thread's `run()` method.
4. If you need COM access from the main thread (e.g., for startup shortcut creation), initialize a separate COM apartment on the main thread. Do not share COM objects between threads.

**Detection:** Intermittent `com_error` with HRESULT `0x8001010E` or `0x80010108`. Crashes that only happen "sometimes" during Explorer autosize.

**Phase relevance:** Module splitting phase and Explorer feature phase.

---

### Pitfall 6: PyInstaller + QWebEngine Hidden Import and Resource Collection Failures

**What goes wrong:** PyInstaller's hooks for PySide6 and QtWebEngine have a history of incomplete collection. The current spec file includes hidden imports for `PySide6.QtWebEngineWidgets`, `PySide6.QtWebEngineCore`, and `PySide6.QtWebChannel`, but several failure modes remain:
- `QtWebEngineProcess.exe` (the Chromium subprocess) is not found at runtime because PyInstaller placed it in the wrong subdirectory relative to the main executable.
- Qt's `resources/` directory (containing `qtwebengine_devtools_resources.pak`, `icudtl.dat`, etc.) is missing or incomplete, causing the WebEngine to show a blank page with no error.
- On PyInstaller version upgrades, hooks change behavior and previously working builds break.

**Why it happens:** QWebEngine is essentially an entire Chromium browser bundled into Qt. It has its own subprocess executable, resource files, and localization data that must be in specific relative paths. PyInstaller's hooks try to handle this but the paths change between Qt versions and across PyInstaller versions.

**Consequences:** The app builds successfully but WebEngine fails at runtime. Blank pages, missing dev tools, or crashes with "Could not find QtWebEngineProcess."

**Prevention:**
1. Pin both PyInstaller and PySide6 versions in `requirements.txt` (or `pyproject.toml`). Do not upgrade one without testing the build.
2. After every PyInstaller build, verify the `dist/Virelo/` directory contains: `QtWebEngineProcess.exe`, the `resources/` folder with `.pak` files, and the `translations/` folder.
3. Add a smoke test to CI: build the app, then run it with a flag that loads the WebEngine and exits (e.g., `--smoke-test` that loads the frontend and checks for the bridge object).
4. Use PyInstaller 6.5+ which has improved QtWebEngine collection for PySide6.

**Detection:** Build from clean checkout, run the resulting executable. If the UI is blank or the log shows "Could not find QtWebEngineProcess," this pitfall was hit.

**Phase relevance:** Build pipeline phase and CI phase.

---

### Pitfall 7: Version String Duplication Across Files

**What goes wrong:** Version strings are duplicated in 5+ locations: `frontend/package.json` (`"1.4.2"`), likely in `app_config.py` or the spec file name, the Inno Setup `.iss` file, and potentially in the app's About dialog or window title. When bumping versions, one or more locations are missed, leading to mismatched versions between the frontend UI, the Python backend, the installer metadata, and the Windows executable properties.

**Why it happens:** The project grew organically. Each component (npm, Python, PyInstaller, Inno Setup) has its own version convention and no single-source-of-truth mechanism was established.

**Consequences:** User confusion ("I installed 1.5 but the About screen says 1.4.2"). Debugging confusion ("which version was this bug report from?"). App store / distribution confusion if versions don't match between the executable properties and the installer.

**Prevention:**
1. Define the version once in `pyproject.toml` under `[project] version = "X.Y.Z"`.
2. Python code reads it via `importlib.metadata.version("virelo")` at runtime, falling back to a hardcoded value for non-installed (dev) runs.
3. The build script reads `pyproject.toml` and injects the version into `frontend/package.json` and the Inno Setup script before building.
4. The PyInstaller spec file reads from `pyproject.toml` as well (spec files are Python, so they can `import tomllib`).

**Detection:** Compare the version shown in the app's UI, the `package.json`, the installer "Add/Remove Programs" entry, and the executable's file properties. If any differ, this pitfall is present.

**Phase relevance:** Build pipeline phase.

---

## Moderate Pitfalls

---

### Pitfall 8: Removing Fake UI Controls Without Updating Frontend State

**What goes wrong:** The React frontend has toggles for auto-update, telemetry, and hidden files that are not wired to the backend. Simply deleting these components from the React code seems safe, but if the frontend's state management (React state, localStorage, or settings loaded from the bridge) still references these keys, the UI breaks or shows console errors. Worse, if the bridge's `save_settings` is called with these now-removed keys by stale frontend code, `SettingsState.apply_partial()` silently drops them (line 63: `if key not in self.KEYS: continue`), which may mask bugs.

**Why it happens:** The bridge was designed to silently ignore unknown keys as a robustness measure, but this also hides the symptom when frontend and backend schemas drift.

**Prevention:**
1. When removing a fake control, also remove it from: the React component tree, any React state/context, any localStorage references, and any hardcoded values in the frontend's bridge call payloads.
2. Change `apply_partial` to log a warning (not silently drop) when it encounters unknown keys: `LOG.warning("Unknown setting key ignored: %s", key)`. This makes drift visible during development.
3. Do the removal in one atomic commit per control: remove the frontend component and verify the settings round-trip still works.

**Detection:** Open browser dev tools in the app. Look for console errors about missing state keys. Check the Python log for unknown keys being silently dropped.

**Phase relevance:** Bridge refactoring phase and frontend cleanup phase.

---

### Pitfall 9: `_resource_path` Duplication Causes Divergent Path Resolution

**What goes wrong:** Two copies of `resource_path` exist (`main.py:198` and `webview.py:33`). The `webview.py` copy has a comment explaining it exists to avoid a circular import. During refactoring, if one copy is updated (e.g., to handle a new resource directory) and the other is not, the app loads resources correctly in one context but fails in another. The frontend might load from the right path while the icon loads from the wrong one, or vice versa.

**Why it happens:** Circular imports between `main.py` and `webview.py` forced the duplication. The circular dependency exists because `main.py` imports `webview.py` for `VireloWebView`, while `webview.py` would need `main.py` for `resource_path`.

**Prevention:**
1. Extract `resource_path` into a standalone utility module (e.g., `paths.py` or add it to `app_config.py`). This module should have zero imports from `main.py`, `webview.py`, or `bridge.py`.
2. Both `main.py` and `webview.py` import from this shared module.
3. This extraction is safe because `resource_path` is a pure function with no Qt dependencies -- it only uses `sys`, `os`, and checks `sys.frozen`.

**Detection:** Search the codebase for `resource_path` and `_resource_path`. If more than one definition exists, this pitfall is active (it currently is).

**Phase relevance:** Module splitting phase. Should be one of the first extractions because it is zero-risk and unblocks other work.

---

### Pitfall 10: Keyboard Hook Lifecycle Race During Key Rebinding

**What goes wrong:** `ShiftSnapRestore.update_binding()` at `main.py:691` unhooks the old key and hooks the new one. The `except Exception: pass` blocks at lines 693-698 silently swallow unhook failures. If `update_binding()` is called rapidly (e.g., user clicks "capture key" twice quickly), the old hook may not be fully removed before the new one is registered, resulting in duplicate hooks. The `keyboard` library uses global state, so two hooks on different keys both fire, causing double-snap behavior.

**Why it happens:** The `keyboard` library's hook/unhook API is not atomic and uses global mutable state. The `except: pass` pattern hides the symptom.

**Prevention:**
1. Add a `threading.Lock` around the entire unhook/hook sequence in `update_binding()`.
2. Replace `except Exception: pass` with `except Exception as e: LOG.warning("Failed to unhook: %s", e)` so failures are visible.
3. Add a guard flag (`self._rebinding = True`) that prevents `_on_press` from firing during rebinding.
4. Consider debouncing `update_binding` calls from the bridge -- ignore calls that arrive within 500ms of each other.

**Detection:** Rapidly change the snap key multiple times via the settings UI. If the app starts responding to both old and new keys, this pitfall is active.

**Phase relevance:** Snap architecture phase.

---

### Pitfall 11: Windows CI Runners Lack Admin Privileges and COM Access

**What goes wrong:** The app requires admin elevation (`is_admin()` check at `main.py:1397`) and COM access for Explorer automation. GitHub Actions `windows-latest` runners:
- Do not run as administrator by default. The `ShellExecuteW("runas", ...)` elevation code will fail silently because there is no interactive UAC prompt.
- May not have COM objects like `Shell.Application` available or properly registered.
- Have limited disk space (~14GB free), and PySide6 + PyInstaller + QWebEngine can consume 1-2GB for the build output.

**Why it happens:** CI runners are non-interactive, resource-constrained environments. Code that assumes an interactive Windows desktop session (UAC prompts, COM automation, system tray) will fail.

**Consequences:** CI builds succeed but smoke tests fail. Integration tests that depend on COM or admin privileges cannot run. The build output may exceed runner disk limits.

**Prevention:**
1. Structure tests in tiers: unit tests (no admin, no COM, no GUI) run in CI; integration tests (COM, admin, GUI) run only locally or on self-hosted runners.
2. For the build job, skip the elevation check and smoke tests that require admin. The CI build proves "it compiles and bundles," not "it runs correctly on a desktop."
3. Monitor disk usage: add `du -sh dist/` after the build step to catch size regressions.
4. Cache PySide6 and node_modules across CI runs to avoid re-downloading ~500MB per build.

**Detection:** CI pipeline passes the build step but fails on any test that touches COM, the system tray, or keyboard hooks.

**Phase relevance:** CI phase.

---

### Pitfall 12: `comtypes` Gen_Py Cache Corruption After Updates

**What goes wrong:** The `comtypes` library generates Python wrappers for COM type libraries and caches them in a `gen_py` directory (typically under `%TEMP%`). This cache can become corrupted after: Windows updates that change COM interface definitions, `comtypes` version upgrades, or Python version changes. When corrupted, all COM operations fail with import errors in the generated wrapper modules.

**Why it happens:** The cache stores auto-generated Python code keyed by COM GUID. If the COM interface changes (even slightly) or the cache is partially written (crash during generation), the stale/corrupt code is imported instead of being regenerated.

**Consequences:** Explorer autosize stops working entirely. The `_ensure_dispatch()` workaround in `main.py:141-162` handles some cases but not all corruption modes.

**Prevention:**
1. On app startup, wrap the initial COM initialization in a try/except that clears the `comtypes` cache on failure and retries: `py -m comtypes.clear_cache` equivalent in Python.
2. In PyInstaller frozen mode, set a custom cache location inside the app's data directory (not `%TEMP%`) so the cache is app-scoped and can be safely deleted.
3. The existing `_ensure_dispatch()` pattern is the right idea -- ensure it covers `comtypes` cache corruption in addition to `win32com` gen_py issues.

**Detection:** After a Windows update, if Explorer autosize stops working and the log shows `ImportError` in comtypes-generated modules, this pitfall was hit.

**Phase relevance:** Explorer feature phase.

---

### Pitfall 13: Deprecated Qt 6 API Calls Generate Warnings and Future Breakage

**What goes wrong:** `main.py:1420-1421` calls `setAttribute(Qt.AA_EnableHighDpiScaling)` and `setAttribute(Qt.AA_UseHighDpiPixmaps)`. These attributes are deprecated in Qt 6 (they are always-on). Currently they generate deprecation warnings in logs. In a future PySide6 version, these attributes may be removed entirely, causing an `AttributeError` crash at startup.

**Why it happens:** The code was ported from a Qt 5 / PySide2 era where these attributes were necessary.

**Prevention:**
1. Remove both `setAttribute` calls. High DPI scaling is always enabled in Qt 6.
2. When upgrading PySide6 in the future, check the Qt deprecation notes and remove any newly-deprecated calls.

**Detection:** Deprecation warnings in the log file mentioning `AA_EnableHighDpiScaling` or `AA_UseHighDpiPixmaps`.

**Phase relevance:** Module splitting phase or build pipeline phase -- trivial fix, can be done anytime.

---

## Minor Pitfalls

---

### Pitfall 14: Debug Logging Left Enabled Globally

**What goes wrong:** `main.py:1228` forces `LOG.setLevel(logging.DEBUG)` when Explorer autosize is enabled, setting the global logger to DEBUG level. This generates high-volume output for all subsystems (not just Explorer), filling log files quickly and slowing the app due to I/O.

**Prevention:** Remove the `LOG.setLevel(logging.DEBUG)` call. The logger is already configured at DEBUG level in `_init_logger()`. If per-subsystem debug logging is needed, use named child loggers (e.g., `logging.getLogger("Virelo.explorer")`).

**Phase relevance:** Module splitting phase.

---

### Pitfall 15: Spec File Named "Windows Toolbox.spec" With Spaces

**What goes wrong:** The spec file `Windows Toolbox.spec` has a space in its name, which causes issues in shell scripts that do not properly quote the path. The `build-installer.ps1` already quotes it, but any future bash/CI script that doesn't will break. Additionally, the stale "Windows Toolbox" name is confusing.

**Prevention:** Rename to `virelo.spec`. Update `build-installer.ps1` and any CI scripts to reference the new name.

**Phase relevance:** Identity/naming phase -- should be one of the first changes.

---

### Pitfall 16: `setQuitOnLastWindowClosed(False)` Without Explicit Quit Path

**What goes wrong:** The app calls `setQuitOnLastWindowClosed(False)` at line 1431 because it is a tray app that should keep running when the window is closed. However, if the tray icon fails to initialize or is not available, there is no way to quit the app except killing the process. The `isSystemTrayAvailable()` check at line 1433 partially mitigates this, but on some Windows configurations (e.g., tablet mode, headless RDP), the tray may report as available but not actually be visible.

**Prevention:** Ensure the quit action is always accessible: via tray icon, via a window close confirmation, or via a timeout that exits if the tray was not successfully shown within N seconds.

**Phase relevance:** App lifecycle phase.

---

### Pitfall 17: Shutdown Cleanup Uses Both `aboutToQuit` and `atexit`

**What goes wrong:** Lines 1440-1443 register cleanup handlers on both `app.aboutToQuit` (Qt signal) and `atexit` (Python). The Qt handler fires during normal `app.exec()` exit. The `atexit` handler fires during Python interpreter shutdown. If both fire, `_stop_background_threads()` and `shift_mgr.cleanup()` are called twice, which may cause errors if the keyboard hook is already unhooked or the thread is already stopped.

**Prevention:** Use only `aboutToQuit` for cleanup. Remove the `atexit` handler. If crash-safety is needed, make the cleanup methods idempotent (safe to call multiple times), which is partially the case due to the `except: pass` blocks but should be made explicit with guard flags.

**Phase relevance:** Module splitting phase.

---

## Phase-Specific Warnings

| Phase Topic | Likely Pitfall | Mitigation |
|-------------|---------------|------------|
| Identity/naming cleanup | Spec file rename breaks build scripts | Search all scripts and docs for "Windows Toolbox" before renaming |
| Build pipeline | Frontend not built before PyInstaller | Add `npm ci && npm run build` as first step; assert `frontend/dist/index.html` exists |
| Build pipeline | Version drift across files | Single source in `pyproject.toml`; build script injects into package.json and .iss |
| WebEngine hardening | Navigation not restricted | Override `acceptNavigationRequest`; disable `LocalContentCanAccessRemoteUrls` in production |
| WebEngine hardening | Bridge method exposure | Audit `channel.objects.bridge` in browser console after changes |
| Bridge refactoring | Silent key dropping masks bugs | Log unknown keys instead of silently ignoring |
| Bridge refactoring | Removing fake controls leaves orphaned state | Remove from React state, localStorage, and bridge payloads atomically |
| Module splitting | Circular imports | Extract pure functions first; Qt classes last; use dependency injection |
| Module splitting | Signal/slot disconnection | Explicit `parent=` on all QObjects; test full lifecycle after each extraction |
| Module splitting | COM apartment crossing | Keep COM init + use + uninit co-located in same thread; document threading contract |
| CI setup | Admin-dependent tests fail | Tier tests: unit (CI) vs integration (local/self-hosted) |
| CI setup | Disk space exceeded | Cache PySide6; monitor `dist/` size; consider artifact cleanup |
| Explorer feature | COM cache corruption | Clear comtypes cache on failure; app-scoped cache directory |
| Snap architecture | Keyboard hook race condition | Lock around unhook/hook; log failures instead of swallowing |

## Sources

- [PyInstaller issue #6387: frozen PySide6.QtWebEngineCore](https://github.com/pyinstaller/pyinstaller/issues/6387) -- HIGH confidence
- [PyInstaller issue #3890: Cannot find QtWebEngineProcess](https://github.com/pyinstaller/pyinstaller/issues/3890) -- HIGH confidence
- [Qt for Python: QWebChannel docs](https://doc.qt.io/qtforpython-6/PySide6/QtWebChannel/QWebChannel.html) -- HIGH confidence
- [Qt for Python: QWebEnginePage.acceptNavigationRequest](https://doc.qt.io/qt-6/qwebenginepage.html) -- HIGH confidence
- [PySide6 memory model and QObject lifetime management](https://forum.qt.io/topic/154590/pyside6-memory-model-and-qobject-lifetime-management) -- MEDIUM confidence
- [PyInstaller changelog: QtWebEngine hook fixes in 6.5+](https://pyinstaller.org/en/v6.5.0/CHANGES.html) -- HIGH confidence
- [Python packaging: single-sourcing version](https://packaging.python.org/en/latest/discussions/single-source-version/) -- HIGH confidence
- [comtypes gen_py cache corruption](https://github.com/enthought/comtypes/issues/182) -- MEDIUM confidence
- [Microsoft: COM apartment threading models](https://learn.microsoft.com/en-us/windows/win32/com/multithreaded-apartments) -- HIGH confidence
- [PySide6 best practices: signal/slot lifetime](https://www.zynu.net/ai-skills/pyside6-best-practices) -- MEDIUM confidence
- Direct codebase analysis of `main.py`, `webview.py`, `bridge.py`, `workers.py`, `settings_state.py`, `build-installer.ps1`, `Windows Toolbox.spec` -- HIGH confidence

---

*Pitfalls audit: 2026-04-24*
