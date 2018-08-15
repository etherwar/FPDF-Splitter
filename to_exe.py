# -*- coding: utf-8 -*-
"""
Date: 8.2.2018
Author: Domen Jurkovic
www.damogranlabs.com

This is a template file to build (freeze) python script to windows executable.
Setup:
1. Copy this script (to_exe.py) and bat file (to_exe.bat) to your project folder
2. Edit to_exe.py. Also, finetune includes & excludes
    Optinally, edit to_exe.bat:
        - change "build" command to bdist_msi or other cx_freeze build options
        - remove pause at the end to close terminal immediatelly
3. Run/re-run to_exe.bat script everytime build mus be updated

Steps this script does:
1. Create build and executable in \build subfolder of to_exe.py script
2. Remove .zip file in \build folder if it already exists (on re-running)
3. Zip all content in \build subfolder

cx_freeze docs:
https://cx-freeze.readthedocs.io/en/latest/
http://cx-freeze.readthedocs.io/en/latest/distutils.html
"""

import os
import sys
import shutil

from cx_Freeze import *

# EDIT according to your project
SCRIPT = "main2.pyw"  # main script to build to .exe
APP_NAME = "ForcePDF"    # also output name of .exe file
DESCRIPTION = "Tool to split and save a selected multipage PDF file as individual, incrementing individual files"
VERSION = "1.0"
GUI = True  # if true, this is GUI based app - no console is displayed
ICON = 'fpdf.ico'  # your icon or None

CREATE_ZIP = True  # set to True if you wish to create a zip once build is generated

# http://msdn.microsoft.com/en-us/library/windows/desktop/aa371847(v=vs.85).aspx
shortcut_table = [
    ("DesktopShortcut",        # Shortcut
     "DesktopFolder",          # Directory_
     "Force PDF Splitter v1.0",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]ForcePDF.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     ),
    ("StartMenuShortcut",  # Shortcut
     "StartMenuFolder",  # Directory_
     "Force PDF Splitter v1.0",  # Name
     "TARGETDIR",  # Component_
     "[TARGETDIR]ForcePDF",  # Target
     None,  # Arguments
     None,  # Description
     None,  # Hotkey
     None,  # Icon
     None,  # IconIndex
     None,  # ShowCmd
     'TARGETDIR'  # WkDir
     )
    ]

# Now create the table dictionary
msi_data = {"Shortcut": shortcut_table}

# Change some default MSI options and specify the use of the above defined tables
bdist_msi_options = {'data': msi_data}

executable_options = {
    'build_exe': {
        # pyqt5 (from official cx_freeze examples)
        'includes': ['wx', 'os', 'PyPDF2'],

        # exclude all other GUIs except Pyqt5
        'excludes': ['PyQt5', 'gtk', 'PyQt4', 'Tkinter'],

        # add your files (like images, ...)
        'include_files': ['force.png'],

        # include msvcr files for building an installer
        'include_msvcr': True,

        # amount of data displayed while freezing
        'silent': [True],

    },

    # Add the msi options for shortcut creation
    'bdist_msi': bdist_msi_options,
}


##############################################################################
# cx_freeze stuff
##############################################################################
_app_name_exe = APP_NAME + ".exe"
if GUI:
    base = 'Win32GUI'
else:
    base = None

# http://cx-freeze.readthedocs.io/en/latest/distutils.html#cx-freeze-executable
exe = Executable(
    script=SCRIPT,
    targetName=_app_name_exe,
    base=base,
    icon=ICON
)

setup(
    name=APP_NAME,
    version=VERSION,
    description=DESCRIPTION,
    options=executable_options,
    executables=[exe]
)
# end of cx_freeze stuff
##############################################################################
print("\n")

cwd = os.getcwd()
build_path = os.path.join(cwd, 'build')

# check if zip already exists, if it does, delete it
zip_name = APP_NAME + ".zip"
zip_path = os.path.join(build_path, zip_name)
if os.path.exists(zip_path):
    os.remove(zip_path)
    print("Zip already exist in \\build subfolder, deleted")

# check if there is anything in \build subfolder
if os.listdir(build_path):  
    print("Distribution created in \\build subfolder")
else:
    print("Empty \\build subfolder, was build successfull?")
    sys.exit(1)

if not CREATE_ZIP:
    print("\nBuild finished.")
    sys.exit(0)

##############################################################################
# distribuition created in \build subfolder, create zip
# create zip file in the root folder
root_zip_path = shutil.make_archive(APP_NAME, 'zip', build_path)
print("Zip created")

# move zip to \build subfolder
shutil.move(root_zip_path, build_path)
print("Zip moved to \\build subfolder")

print("\nBuild & zip finished.")
sys.exit(0)
