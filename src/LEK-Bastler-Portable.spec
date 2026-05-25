# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

try:
    SPEC_DIR = Path(__file__).resolve().parent
except NameError:
    cwd = Path.cwd()
    spec_in_src = cwd / 'src' / 'LEK-Bastler-Portable.spec'
    SPEC_DIR = spec_in_src.parent if spec_in_src.exists() else cwd

PROJECT_ROOT = SPEC_DIR.parent
SRC_DIR = PROJECT_ROOT / 'src'

a = Analysis(
    [str(SRC_DIR / 'main.py')],
    pathex=[str(SRC_DIR), str(PROJECT_ROOT)],
    binaries=[],
    datas=[
        (str(PROJECT_ROOT / 'data'), 'data'),
        (str(SRC_DIR / 'app_icon.ico'), '.'),
    ],
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
    name='LEK-Bastler-Portable',
    icon=str(SRC_DIR / 'app_icon.ico'),
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
    version=str(SPEC_DIR / 'build_version_info.txt'),
)
