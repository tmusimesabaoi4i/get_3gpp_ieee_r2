# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main_fetch_and_convert.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=['win32timezone', 'requests', 'urllib3', 'certifi', 'idna', 'charset_normalizer', 'pandas', 'numpy', 'openpyxl'],
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
    a.binaries,
    a.datas,
    [],
    name='mytool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
