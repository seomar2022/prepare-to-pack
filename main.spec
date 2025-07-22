# -*- mode: python ; coding: utf-8 -*-

block_cipher = None
from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo, FixedFileInfo
)

version = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(2, 1, 0, 0),
        prodvers=(2, 1, 0, 0),
        mask=0x3f,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(2025, 722)
    ),
    kids=[]
)

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('settings/header.csv', 'settings'),
        ('resources/img/document.png', 'resources/img'),
        ('resources/img/favicon.ico', 'resources/img'),
        ('resources/img/gear.png', 'resources/img'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='PrepareToPack',
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
    icon='resources/img/favicon.ico',
    version=version,
)
