# -*- mode: python ; coding: utf-8 -*-


import os

tcl_dll = r"C:\ProgramData\miniconda3\Library\bin\tcl86t.dll"
tk_dll = r"C:\ProgramData\miniconda3\Library\bin\tk86t.dll"

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[(tcl_dll, '.'), (tk_dll, '.'),
              (r'C:\ProgramData\miniconda3\Library\bin\ffi.dll', '.'),
              (r'C:\ProgramData\miniconda3\Library\bin\liblzma.dll', '.'),
              (r'C:\ProgramData\miniconda3\Library\bin\libbz2.dll', '.'),
              (r'C:\ProgramData\miniconda3\Library\bin\libexpat.dll', '.'),
              (r'C:\ProgramData\miniconda3\Library\bin\sqlite3.dll', '.')],
    datas=[('icon.ico', '.'), ('阅读理解.xlsx', '.')],
    hiddenimports=['excel_reader', 'question', 'word_generator', 'pandas', 'openpyxl', 'openpyxl.cell', 'openpyxl.reader.excel', 'pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.nattype', 'pandas._libs.tslibs.parsing', 'pandas._libs.tslibs.timezones', 'pandas._libs.tslibs.offsets', 'pandas._libs.tslibs.strptime'],
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
    name='考试试卷生成器',
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
    icon=['icon.ico'],
)
