# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path


python_root = Path(sys.executable).resolve().parent
tcl_root = python_root / "tcl"
tkinter_root = python_root / "Lib" / "tkinter"
tk_binaries = []
for dll_name in ("tcl86t.dll", "tk86t.dll"):
    dll_path = python_root / "DLLs" / dll_name
    if dll_path.exists():
        tk_binaries.append((str(dll_path), "."))

tk_datas = []
if tcl_root.exists():
    tcl_library = tcl_root / "tcl8.6"
    tk_library = tcl_root / "tk8.6"
    if tcl_library.exists():
        tk_datas.append((str(tcl_library), "_tcl_data"))
    if tk_library.exists():
        tk_datas.append((str(tk_library), "_tk_data"))
if tkinter_root.exists():
    tk_datas.append((str(tkinter_root), "tkinter"))


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=tk_binaries,
    datas=tk_datas,
    hiddenimports=['tkinter', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['tk_runtime_hook.py'],
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
    name='JIIN_DNC_Manager_V2',
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
