# -*- coding: utf-8 -*-
import sys
import os
import ctypes
from PySide6.QtWidgets import QApplication

# No system path hacking needed if run via run.py as module

from .gui.main_window import RPAWindow
from .utils import global_exception_handler

# ---------------------------------------------------------
# Core setup
# ---------------------------------------------------------
def setup_env():
    try:
        os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
        os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
        ctypes.windll.shcore.SetProcessDpiAwareness(1) 
    except:
        try: ctypes.windll.user32.SetProcessDPIAware()
        except: pass

def main():
    setup_env()
    sys.excepthook = global_exception_handler
    
    app = QApplication(sys.argv)
    win = RPAWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
