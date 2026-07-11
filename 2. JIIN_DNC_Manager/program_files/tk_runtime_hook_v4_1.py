import os
import shutil
import sys
from pathlib import Path


base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
app_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

bundled_tcl = base_dir / "_tcl_data"
bundled_tk = base_dir / "_tk_data"
runtime_root = app_dir / "data_v4_1" / "tk_runtime"
runtime_tcl = runtime_root / "tcl8.6"
runtime_tk = runtime_root / "tk8.6"


def copy_runtime_tree(source: Path, target: Path, marker_name: str) -> Path:
    if not source.exists():
        return source
    try:
        if not (target / marker_name).exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(source, target, dirs_exist_ok=True)
    except Exception:
        return source
    return target if target.exists() else source


tcl_library = copy_runtime_tree(bundled_tcl, runtime_tcl, "init.tcl")
tk_library = copy_runtime_tree(bundled_tk, runtime_tk, "tk.tcl")

if str(base_dir) not in sys.path:
    sys.path.insert(0, str(base_dir))

if tcl_library.exists():
    os.environ["TCL_LIBRARY"] = str(tcl_library)
if tk_library.exists():
    os.environ["TK_LIBRARY"] = str(tk_library)
