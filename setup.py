import cx_Freeze
import sys
import os
base = None

if sys.platform == 'win32':
    base = "Win32GUI"


os.environ['TCL_LIBRARY']=r"E:\2019\Python\Portable Python-3.7.4 x64\App\Python\tcl\tcl8.6"
os.environ['TK_LIBRARY']=r"E:\2019\Python\Portable Python-3.7.4 x64\App\Python\tcl\tk8.6"

executables = [cx_Freeze.Executable("Site Risk Assessment Tool.py", base=base,icon="site.ico")]


cx_Freeze.setup(
    name = "SRA Tool",
    options = {"build_exe": {"packages":["tkinter","os","csv","time","datetime","openpyxl","pandas","matplotlib","xlrd","DataFrame","numpy"], "include_files":['tcl86t.dll','tk86t.dll', 'icons2']}},
    version = "4.0",
    description = "PCPandey Application",
    executables = executables
    )
