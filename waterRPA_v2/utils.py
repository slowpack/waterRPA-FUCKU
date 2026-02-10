# -*- coding: utf-8 -*-
import sys
import os
import time
import traceback
from .config import GLOBAL_CONFIG

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def get_log_path():
    # Log path logic might need adjustment if we are in a subdirectory now (waterRPA_v2)
    # But usually we want it relative to the entry point or user home.
    # The original code used __file__ which was waterRPA.py in the root.
    # Now this file is in waterRPA_v2/utils.py.
    # If we want the log in the project root (parent of waterRPA_v2), we should go up one level.
    # However, let's keep it simple and relative to the package for now, or use the parent of `main.py`.
    # `get_base_dir` will return `.../waterRPA_v2` if run from there? No, `__file__` is `.../waterRPA_v2/utils.py`.
    # `os.path.dirname` is `.../waterRPA_v2`.
    
    # Let's adjust get_base_dir to return the parent directory if we are not frozen, 
    # assuming we want the log file in the same place as before (project root).
    base = os.path.dirname(os.path.abspath(__file__)) # waterRPA_v2
    project_root = os.path.dirname(base) # waterRPA-FUCKU
    
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), "rpa_debug_log.txt")
    else:
        return os.path.join(project_root, "rpa_debug_log.txt")

def write_log(msg):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {msg}"
    if GLOBAL_CONFIG["log_to_file"]:
        try:
            with open(get_log_path(), "a", encoding="utf-8") as f:
                f.write(formatted_msg + "\n")
        except: pass

def global_exception_handler(exctype, value, tb):
    err_msg = "".join(traceback.format_exception(exctype, value, tb))
    write_log(f"!!! 严重崩溃 !!! {value}\n{err_msg}")
    sys.__excepthook__(exctype, value, tb)
