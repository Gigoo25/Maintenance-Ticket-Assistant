@echo off

REM Minimize Window
REM if not "%1" == "min" start /MIN cmd /c %0 min & exit/b

REM Set location of directory
%~d0
cd %~dp0
set root=%cd%
set script_location=%root%\script.py

REM Update files
call Tools\Functions\Update_function

REM Install prerequisites
pip install -r %root%\Tools\Requirements.txt

REM Clear screen
CLS

REM Run script
python %script_location%