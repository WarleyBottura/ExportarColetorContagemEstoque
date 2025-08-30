# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Exportador_contagem_Ver2.1.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'pandas.tests', 'numpy.random._examples', 'numpy.testing', 'numba', 'llvmlite', 'tbb', 'tbb4py'],
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
    name='ExportadorContagem_app_V2.1',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=['libssl-3.dll', 'libcrypto-3.dll', 'ucrtbase.dll', 'VCRUNTIME140.dll', 'VCRUNTIME140_1.dll'],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icone.ico'],
)
