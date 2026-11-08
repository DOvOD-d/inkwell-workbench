# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

pdfium_datas = collect_data_files('pypdfium2_raw')
pdfium_bins = collect_dynamic_libs('pypdfium2_raw')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=pdfium_bins,
    datas=[('assets', 'assets')] + pdfium_datas,
    hiddenimports=['pythoncom', 'pywintypes', 'win32com', 'win32com.client',
                   'PIL', 'PIL.Image', 'pypdfium2', 'pypdfium2_raw'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['QtWebEngine', 'tkinter'],
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
    name='Word模板工作台',
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
    icon=['assets\\icon.ico'],
)
