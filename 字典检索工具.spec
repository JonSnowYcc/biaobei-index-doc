# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['index_v3_20250913.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['xlwings', 'docx', 'charset_normalizer', 'portalocker', 'win32com', 'docx.oxml.ns', 'docx.enum.text', 'docx.shared', 'docx.oxml', 'docx.enum', 'docx.oxml.shared', 'win32timezone', 'pythoncom', 'pywintypes', 'win32api', 'win32con'],
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
    name='字典检索工具',
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
)
