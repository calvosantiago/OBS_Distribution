# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
from pathlib import Path
import country_converter as coco

# ======================
#  DATOS A EMPAQUETAR
# ======================

datas = []

# --- 1) Datos internos de country_converter (evita FileNotFoundError) ---
# Hook estándar:
datas += collect_data_files('country_converter', includes=['country_data.tsv'])
# Redundancia explícita (por si el hook falla en tu entorno):
ccdir = Path(coco.__file__).parent
datas.append((str(ccdir / 'country_data.tsv'), 'country_converter'))

# --- 2) Solver CBC de PuLP (si usas PuLP.CBC) ---
# Incluye toda la carpeta solverdir (cbc.exe y demás)
datas += collect_data_files('pulp', includes=['solverdir/*'])

# --- 3) Tus archivos locales (opcional; ajusta rutas si hace falta) ---
def add_data(src: str, dest: str):
    p = Path(src)
    if p.exists():
        datas.append((str(p), dest))
    else:
        print(f"⚠️ WARNING: archivo no encontrado y no se empaqueta: {p}")

# Ejemplos (coméntalos o ajusta a tu caso):
# add_data('SUDOKU.xlsx', '.')                    # raíz del ejecutable
# add_data('input/Areas_Paises.xlsx', 'input')    # subcarpeta 'input'
# add_data('config/Parametros.xlsx', 'config')

# ======================
#   CONFIGURACIÓN BASE
# ======================

a = Analysis(
    ['Distribucion_OBS.py'],
    pathex=[],
    binaries=[],
    datas=datas,
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
    name='Distribucion_OBS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,   # ← ponlo en False cuando ya funcione todo
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
