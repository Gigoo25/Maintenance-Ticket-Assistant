@echo off

REM Set location of directory
%~d0
cd %~dp0
set root=%cd%
set script_location=%root%\script.py

REM Update files
call Tools\Functions\Update_function

REM Install prerequisites
pip install -r %Output%\Tools\Requirements.txt

REM Clear screen
CLS

REM Run script
python %script_location%

REM Pause at the end for testing purposes
pause