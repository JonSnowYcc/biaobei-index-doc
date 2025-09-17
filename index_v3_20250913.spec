# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.building.build_main import Analysis, PYZ, EXE, COLLECT

a = Analysis(
    ['index_v3_20250913.py'],
    pathex=[],
    binaries=[],
    datas=collect_data_files('xlwings'),
    hiddenimports=[
        'xlwings',
        'docx',
        'charset_normalizer',
        'portalocker',
        'win32com',
        'docx.oxml.ns',
        'docx.enum.text',
        'docx.shared',
        'docx.oxml',
        'docx.enum',
        'docx.oxml.shared',
        'win32timezone',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32con',
        'tkinter',
        'tkinter.filedialog',
        'json',
        're',
        'os',
        'threading',
        'datetime',
        'random'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False
)

# datas 已在 Analysis 中通过 collect_data_files('xlwings') 注入

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='字典检索工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为False以隐藏控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None  # 如果有图标文件可以在这里指定
)
