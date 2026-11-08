# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build specification for DTV YAML Generator.

Build a single-file executable that bundles the Python script together with
the two data files it needs at runtime:

    * BroadcastTransmitter_dtv_simplified_merged.json  (dataset)
    * callsign_name_lookup.json                        (callsign lookup table)

Usage (from the project root):

    pyinstaller --clean --noconfirm dtv_yaml_generator.spec

The resulting executable is written to ``dist/``:

    * Windows: ``dist/DTV_YAML_Generator.exe``
    * macOS:   ``dist/DTV_YAML_Generator``
"""

import os

block_cipher = None

# Data files bundled into the executable. They are unpacked next to the
# script at runtime (``sys._MEIPASS``) and located by ``_resource_dirs()``.
datas = [
    ("BroadcastTransmitter_dtv_simplified_merged.json", "."),
    ("callsign_name_lookup.json", "."),
]

a = Analysis(
    ["dtv_yaml_generator.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="DTV_YAML_Generator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
