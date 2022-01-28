from cx_Freeze import setup, Executable
import shutil
import os

try:
    shutil.rmtree('build')
    shutil.rmtree('dist')
except:
    pass
base = 'Win32GUI'
# base = None

executables = [Executable('main.py',
                          target_name='BOM_check.exe',
                          base=base,
                          icon='./icon.ico'),
               ]


# packages = ["main"]
options = {
    'build_exe': {
        'include_msvcr': True,
    }
}

setup(
    name="Orion config",
    options=options,
    version=1.1,
    description='',
    executables=executables
)