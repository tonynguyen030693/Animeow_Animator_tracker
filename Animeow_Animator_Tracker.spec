# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['E:\\Tool_tinh_diem\\Tool_UI_update\\Animeow_Animator_Tracker.py'],
    pathex=[],
    binaries=[],
    datas=[('E:\\Tool_tinh_diem\\Tool_UI_update\\Animeow_logo.jpg', '.'), ('E:\\Tool_tinh_diem\\Tool_UI_update\\Animeow_logo.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Animeow_Animator_Tracker',
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
    icon='E:\\Tool_tinh_diem\\Tool_UI_update\\Animeow_logo.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Animeow_Animator_Tracker',
)
