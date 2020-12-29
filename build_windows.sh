#!/bin/bash
# first make sure your python is built by enabling share
# e.g. CONFIGURE_OPTS=--enable-shared pyenv install 3.9.1
# command to build executable file in Windows
# pyinstaller only works on version 3.7 on windows
# pyinstaller --onefile --windowed --noupx  .\excel_to_csv.py