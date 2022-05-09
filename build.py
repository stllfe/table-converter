import base64

from distutils.sysconfig import get_python_lib
from pathlib import Path

from PyInstaller.__main__ import run

# include all the pywin32 DLLs automatically
dll_path = Path(get_python_lib()) / 'pywin32_system32'
binaries = tuple(dll_path.glob('*.dll'))


# hacky fix for bitmap icon loading with pyinstaller packaging
# https://stackoverflow.com/questions/9929479/embed-icon-in-python-script
with open('src/icon.py', 'w') as icon_module:
    icon_module.write("img = '")
with open('icon.ico', 'rb') as icon, open('src/icon.py', 'ab+') as icon_module:
    b64str = base64.b64encode(icon.read())
    icon_module.write(b64str)
with open('src/icon.py', 'a') as icon_module:
    icon_module.write("'")


run([
    'main.py',
    '--onefile',
    '--windowed',
    '--icon=icon.ico',
    *(f'--add-binary={path};.' for path in binaries),
])
