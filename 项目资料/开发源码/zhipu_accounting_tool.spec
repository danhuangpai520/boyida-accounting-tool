# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['zhipu_accounting_app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\boyida_truck.png', 'assets'), ('assets\\boyida_truck.ico', 'assets'), ('assets\\jingzhe_header_line.png', 'assets')],
    hiddenimports=[],
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
    name='zhipu_accounting_tool',
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
    icon=['assets\\boyida_truck.ico'],
)
