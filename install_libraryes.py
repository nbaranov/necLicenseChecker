import subprocess
import sys


def install_and_import(package):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])


libs = "openpyxl", "PySimpleGUI", "requests", "bs4", "lxml", 

for lib in libs:
    install_and_import(lib)
