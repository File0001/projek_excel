# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build spec for CheapNFriendly Excel to PDF Converter.

Produces a single windowed executable (dist/ExcelToPdfConverter.exe) that
bundles all Python dependencies, the tkdnd drag-and-drop binaries, the
CustomTkinter theme resources and the application icon.
"""

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

projectRoot = os.path.abspath(os.getcwd())

# Bundle runtime data files (tkdnd binaries/tcl, customtkinter assets, icon)
datas = []
datas += collect_data_files("tkinterdnd2")
datas += collect_data_files("customtkinter")
datas += [(os.path.join(projectRoot, "icon.ico"), ".")]

# Ensure dynamically-imported submodules are picked up
hiddenimports = []
hiddenimports += collect_submodules("customtkinter")
hiddenimports += ["tkinterdnd2", "openpyxl", "pypdf", "win32com.client", "pythoncom"]


a = Analysis(
    ["run.py"],
    pathex=[projectRoot],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "PyQt5", "PySide2", "PySide6"],
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
    name="ExcelToPdfConverter",
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
    icon=os.path.join(projectRoot, "icon.ico"),
)
