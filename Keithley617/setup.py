import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [Executable('JVMeasurementApp.py', base=base)]

setup(
    name='JVMeasurementApp',
    version='1.0',
    description='JV Measurement Application',
    executables=executables,
    options={   
        'build_exe': {
            'packages': ['numpy', 'matplotlib', 'tkinter', 'serial'],
            'include_files': ['Keithley617.py']  # Adjust as necessary
        }
    }
)
