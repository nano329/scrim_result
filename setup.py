from cx_Freeze import setup, Executable
import sys

buildOptions = dict(packages= ['PyQt5', 'main', 'meta', 'os', 'copy', 'time', 'datetime', 'openpyxl', 'xlsxwriter', 'KartRider', 'collections'], excludes= [])
exe = [Executable("./main_cmd.py", base= "Win32GUI")]

setup(
    name = 'test',
    version = '0.1',
    author = "test",
    description = "test",
    options = dict(build_exe = buildOptions),
    executables = exe
)