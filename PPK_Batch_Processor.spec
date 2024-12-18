# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['PPK batch processor.py'],
    pathex=[],
    binaries=[],
    datas=[('tools.py', '.'), ('Drone_GNSS_app_v1.3.py', '.'), ('assets/*', 'assets/'), ('rtklib/*', 'rtklib/')],
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
    [],
    exclude_binaries=True,
    name='PPK_Batch_Processor',
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
    icon=['assets\\icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PPK_Batch_Processor',
)
