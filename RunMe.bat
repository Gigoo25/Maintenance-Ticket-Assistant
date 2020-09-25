@echo off

REM Set location of directory
%~d0
cd %~dp0
set root=%cd%
set script_dir=%root%\script.py
REM Install prerequisites
pip install -U selenium
pip install -U pynput
pip install -U pathlib
pip install -U openpyxl
pip install -U Pillow
pip install -U PySimpleGUI
REM Clear
CLS
REM Run script
python %script_dir%
pause