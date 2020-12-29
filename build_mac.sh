#!/bin/bash
# first make sure your python is built by enabling share
# e.g. CONFIGURE_OPTS=--enable-shared pyenv install 3.9.1
# command to build executable file in MacOS
pyinstaller --onefile --add-binary='/System/Library/Frameworks/Tk.framework/Tk':'tk'  --add-binary='/System/Library/Frameworks/Tcl.framework/Tcl':'tcl' excel_to_csv.py