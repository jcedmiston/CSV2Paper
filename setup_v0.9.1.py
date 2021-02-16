import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"build_exe": "./dist/CSV2Paper_0.9.1", "include_files": [('C:/Users/Joseph/Developer/CSV2PAPER/BundledResources', './BundledResources')], "packages": ["win32com.client"], "silent": True}

# GUI applications require a different base on Windows (the default is for
# a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

target = Executable(
    script="GUI.py",
    target_name="CSV2Paper.exe",
    base=base
    )
    #icon="icon.ico"

setup(  name = "CSV2Paper",
        version = "0.9.1",
        description = "Convert CSV rows to individual pages",
        options = {"build_exe": build_exe_options},
        executables = [target])