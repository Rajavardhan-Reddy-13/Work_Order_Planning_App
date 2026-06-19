# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for WO_Planning.py
# Run:  pyinstaller WO_Planning.spec

block_cipher = None

a = Analysis(
    ['WO_Planning.py'],
    pathex=[],
    binaries=[],
    datas=[],           # WO_Planning_App.xlsx stays NEXT TO the .exe, not bundled inside
    hiddenimports=[
        'pyodbc',
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'numpy',
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'PIL', 'tkinter',
        'IPython', 'jupyter', 'notebook',
    ],
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
    name='WO_Planning',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,              # add an .ico path here if you have one
)
