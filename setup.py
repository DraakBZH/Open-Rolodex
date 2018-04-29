import sys
import os
from cx_Freeze import setup, Executable
os.environ['TCL_LIBRARY'] = r'C:\Users\Draak\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\Draak\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'

# Dependencies are automatically detected, but it might need fine tuning.

build_exe_options = { "packages": ["os"]}

# GUI applications require a different base on Windows (the default is for a
# console application).

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(name="Open Rolodex", 
    version="0.1.0.0", 
    description="Open Rolodex",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)])
