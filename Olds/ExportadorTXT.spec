# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['exportador_txt_concatenacao_dinamica_a_partir_do_xls_gui_tkinter2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'seaborn', 'scipy', 'pyarrow', 'fastparquet', 'tables', 'sqlalchemy', 'lxml', 'html5lib', 'bs4', 'jinja2', 'google.cloud.bigquery', 'pandas.tests', 'pandas._testing', 'pandas.plotting', 'pandas.io.formats', 'pandas.io.sas', 'pandas.io.gbq', 'pandas.io.json', 'pandas.io.orc', 'pandas.io.sql', 'pandas.io.stata', 'pandas.io.feather', 'pandas.io.parquet', 'pandas.io.clipboards', 'pandas.io.pickle', 'numpy.random._examples', 'numpy.random._pickle', 'numpy.f2py', 'numpy.distutils', 'numpy.testing'],
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
    name='ExportadorTXT',
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
    version='verinfo.txt',
    icon=['icone.ico'],
)
