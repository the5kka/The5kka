import os
import sys
from pathlib import Path


base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
tcl_library = base_dir / "_tcl_data"
tk_library = base_dir / "_tk_data"

if str(base_dir) not in sys.path:
    sys.path.insert(0, str(base_dir))

if tcl_library.exists():
    os.environ.setdefault("TCL_LIBRARY", str(tcl_library))
if tk_library.exists():
    os.environ.setdefault("TK_LIBRARY", str(tk_library))
