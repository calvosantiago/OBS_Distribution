# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules

datas = []
hiddenimports = []
datas += collect_data_files('country_converter')
hiddenimports += collect_submodules('country_converter')


a = Analysis(
    ['..\\distribucion_OBS.py'],
    pathex=[],
    binaries=[('C:\\Users\\uscp9a\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\pulp\\solverdir\\cbc\\win\\i64\\cbc.exe', 'pulp/solverdir/cbc/win/i64')],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='distribucion_OBS',
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
