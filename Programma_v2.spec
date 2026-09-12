# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Programma_v2.py'],
    pathex=[],
    binaries=[],
    datas=[('McCATMuS_nfd_nofix_V1.mlmodel', '.')],
    hiddenimports=['kraken', 'kraken.pageseg', 'kraken.rpred', 'kraken.lib.models', 'groq', 'pydantic', 'pandas', 'openpyxl', 'PIL', 'customtkinter'],
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
    name='Programma_v2',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Programma_v2',
)
