"""Entry point for STU-MIS Backup Tool."""
import sys
import os
import site
import subprocess

# Ensure src/ is on the path when running as a script or frozen exe
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Ensure user site-packages is available (needed when launched via .bat on Windows)
for _p in site.getsitepackages() + [site.getusersitepackages()]:
    if _p not in sys.path:
        sys.path.insert(0, _p)

# Auto-install required packages if missing
_REQUIRED = ["xlsxwriter", "customtkinter", "requests", "bs4", "lxml"]
for _pkg in _REQUIRED:
    try:
        __import__(_pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", _pkg])

from gui import App

if __name__ == "__main__":
    app = App()
    app.mainloop()
