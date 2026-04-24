# -*- mode: python ; coding: utf-8 -*-

import re
from pathlib import Path

# Parse APP_VERSION via regex -- do NOT import app_config directly.
# Importing app_config in spec context may trigger PySide6 import chain.
_cfg = Path("app_config.py").read_text()
_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', _cfg)
APP_VERSION = _match.group(1) if _match else "0.0.0"

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ("icon.ico", "."),
        ("frontend/dist", "frontend/dist"),
    ],
    hiddenimports=[
        'bridge', 'webview', 'settings_state', 'snap_service',
        'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineCore', 'PySide6.QtWebChannel',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Virelo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Virelo',
)
