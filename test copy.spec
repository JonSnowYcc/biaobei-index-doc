# -*- mode: python ; coding: utf-8 -*-
import os
import sys
python_dll = os.path.join(sys.base_prefix, 'python311.dll')  # 根据您的Python版本调整文件名

block_cipher = None

a = Analysis(
    ['test copy.py'],
    pathex=[],
    binaries=[
        (python_dll, '.'),  # 添加Python DLL
        ('C:\\Windows\\System32\\vcruntime140.dll', '.'),  # 添加运行时依赖
        ('C:\\Windows\\System32\\msvcp140.dll', '.'),      # 添加运行时依赖
    ],
    datas=[],
    hiddenimports=[
        'win32api', 
        'win32con',
        'win32com',
        'win32com.client',
        'win32timezone',
        'pythoncom',
        'pywintypes',
        'xlwings',
        'xlwings.constants'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 改回 False 以隐藏控制台
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
