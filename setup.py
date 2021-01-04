import os
import sys
from cx_Freeze import setup, Executable


PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

syspath = r"C:\Users\[username]\AppData\Local\Programs\Python\Python36-32/DLLs"

buildOptions = dict(
    packages=[],
    excludes=[]
)

executables = [
    Executable('program.py', base=base)
]

setup(name='Programa clientes',
      version='1.0',
      options=dict(build_exe=buildOptions),
      description='Gerador de word com base em cliente.',
      executables=executables
)