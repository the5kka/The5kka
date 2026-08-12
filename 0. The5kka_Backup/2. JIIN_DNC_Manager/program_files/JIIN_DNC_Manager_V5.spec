# -*- mode: python ; coding: utf-8 -*-

import os
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

spec_root = Path(SPECPATH).resolve()
project_root = spec_root.parent
asset_datas = []
simtek_logo = project_root / "data" / "simtek_logo.png"
if simtek_logo.exists():
    asset_datas.append((str(simtek_logo), "."))
sps_rule_file = spec_root / "data_v5" / "sps_rules_default.xlsx"
if sps_rule_file.exists():
    asset_datas.append((str(sps_rule_file), "."))
sps_catalog_file = spec_root / "data_v5" / "sps_program_catalog.json"
if sps_catalog_file.exists():
    asset_datas.append((str(sps_catalog_file), "."))


a = Analysis(
    ['main_v5.py'],
    pathex=[],
    binaries=tk_binaries,
    datas=tk_datas + asset_datas,
    hiddenimports=['tkinter', '_tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['tk_runtime_hook_v5.py'],
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
    name='JIIN_DNC_Manager_V5',
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
