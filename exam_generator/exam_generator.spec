# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 收集tkinter相关数据和模块
tkinter_datas = collect_data_files('tkinter')
tkinter_modules = collect_submodules('tkinter')

# 获取Python安装目录
python_dir = os.path.dirname(sys.executable)

# 添加tkinter所需的DLL文件
tkinter_dlls = []
if sys.platform.startswith('win'):
    # Windows平台需要包含tkinter相关的DLL文件
    # 使用找到的DLL文件路径（使用原始字符串避免转义问题）
    tcl_dll = r"C:\Program Files\Autodesk\Maya2025\Python\DLLs\tcl86t.dll"
    tk_dll = r"C:\Program Files\Autodesk\Maya2025\Python\DLLs\tk86t.dll"
    if os.path.exists(tcl_dll):
        tkinter_dlls.append((tcl_dll, '.'))
    if os.path.exists(tk_dll):
        tkinter_dlls.append((tk_dll, '.'))
        
    # 添加其他可能需要的系统DLL文件
    system_dlls = [
        "liblzma.dll",
        "LIBBZ2.dll",
        "libexpat.dll",
        "ffi.dll",
        "sqlite3.dll"
    ]
    
    dll_dir = r"C:\Program Files\Autodesk\Maya2025\Python\DLLs"
    for dll_name in system_dlls:
        dll_path = os.path.join(dll_dir, dll_name)
        if os.path.exists(dll_path):
            tkinter_dlls.append((dll_path, '.'))

# 创建Analysis对象
a = Analysis(
    ['main.py'],
    pathex=[os.getcwd()],
    binaries=tkinter_dlls,
    datas=tkinter_datas,
    hiddenimports=tkinter_modules + ['openpyxl', 'pandas._libs.tslibs.nattype'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 创建PYZ对象
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 创建EXE对象
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='考试试卷生成器',
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 改为True以便查看错误信息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
