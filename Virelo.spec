# -*- mode: python ; coding: utf-8 -*-

import re
from pathlib import Path

from PyInstaller.config import CONF

# Parse product constants without importing application modules. Importing the
# package in the spec context can trigger the PySide6 import chain.
_config_text = Path("virelo/app/config.py").read_text(encoding="utf-8")


def _read_string_constant(name):
    match = re.search(rf'^{name}\s*=\s*"([^"]+)"\s*$', _config_text, re.MULTILINE)
    if not match:
        raise RuntimeError(f"{name} was not found in virelo/app/config.py.")
    return match.group(1)


APP_NAME = _read_string_constant("APP_NAME")
APP_VERSION = _read_string_constant("APP_VERSION")
APP_PUBLISHER = _read_string_constant("APP_PUBLISHER")
APP_SUPPORT_URL = _read_string_constant("APP_SUPPORT_URL")

_version_match = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", APP_VERSION)
if not _version_match:
    raise RuntimeError("APP_VERSION must contain exactly three numeric components.")
_version_tuple = tuple(int(part) for part in _version_match.groups()) + (0,)

# PyInstaller accepts a version-resource description file. Generate it in the
# work directory so the executable metadata always derives from APP_VERSION.
_version_resource = Path(CONF["workpath"]) / "Virelo-version-info.txt"
_version_resource.parent.mkdir(parents=True, exist_ok=True)
_version_resource.write_text(
    f"""VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={_version_tuple!r},
    prodvers={_version_tuple!r},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [
          StringStruct('CompanyName', {APP_PUBLISHER!r}),
          StringStruct('FileDescription', {APP_NAME!r}),
          StringStruct('FileVersion', {APP_VERSION!r}),
          StringStruct('InternalName', {APP_NAME!r}),
          StringStruct('OriginalFilename', 'Virelo.exe'),
          StringStruct('ProductName', {APP_NAME!r}),
          StringStruct('ProductVersion', {APP_VERSION!r}),
          StringStruct('Comments', {APP_SUPPORT_URL!r})
        ]
      )
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
""",
    encoding="utf-8",
)

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("LICENSE", "."),
        ("icon.ico", "."),
        ("frontend/dist", "frontend/dist"),
    ],
    hiddenimports=[
        "virelo",
        "virelo.app",
        "virelo.app.window",
        "virelo.app.config",
        "virelo.app.webview",
        "virelo.app.__main__",
        "virelo.bridge",
        "virelo.bridge.bridge",
        "virelo.bridge.capture_guard",
        "virelo.services",
        "virelo.services.snap",
        "virelo.platform.theme",
        "virelo.platform.startup",
        "virelo.services.explorer_columns",
        "virelo.settings",
        "virelo.settings.persistence",
        "virelo.settings.state",
        "virelo.workers",
        "virelo.workers.key_capture",
        "virelo.workers.explorer",
        "virelo.platform",
        "virelo.platform.win32_helpers",
        "virelo.platform.resources",
        "virelo.platform.paths",
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebChannel",
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
    name="Virelo",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    version=str(_version_resource),
    codesign_identity=None,
    entitlements_file=None,
    icon=["icon.ico"],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Virelo",
)
