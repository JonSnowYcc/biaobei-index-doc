from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.building.build_main import Analysis, PYZ, EXE

a = Analysis(
    ['test copy.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'xlwings',
        'docx',
        'charset_normalizer',
        'portalocker',
        'win32com',
        'python-docx',
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
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False
)

# 添加 xlwings 相关文件
xlwings_datas = collect_data_files('xlwings')
a.datas.extend(xlwings_datas)

pyz = PYZ(a.pure, a.zipped_data)

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
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app.ico'  # 如果您有图标文件的话
) 