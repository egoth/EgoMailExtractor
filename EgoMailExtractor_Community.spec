# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

root = Path(SPECPATH)
icon = str(root / "assets" / "egomail_icon.ico")

a = Analysis(
    [str(root / "main.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[
        (str(root / "help"), "help"),
        (str(root / "assets"), "assets"),
        (str(root / "LICENSE.txt"), "."),
        (str(root / "NOTICE"), "."),
        (str(root / "RELEASE_1.39_COMMUNITY.txt"), "."),
        (str(root / "EDITION.txt"), "."),
    ],
    hiddenimports=['pythoncom','pywintypes','win32com','win32com.client','win32timezone'],
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
    name="EgoMailExtractor_Community",
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
    icon=icon,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="EgoMailExtractor_Community",
)
