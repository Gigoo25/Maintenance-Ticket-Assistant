@echo off

REM Minimize Window
if not "%1" == "min" start /MIN cmd /c %0 min & exit/b

REM Set location of directory
%~d0
cd %~dp0
set root=%cd%
set script_location=%root%\script.py

REM Update files
call Tools\Functions\Update_function

REM Install prerequisites
<<<<<<< HEAD
pip install -r %root%\Tools\Requirements.txt
=======
pip install -r %Output%\Tools\Requirements.txt
>>>>>>> f184b098eed66b918e9cee3b431c4a5248cd4ec9

REM Clear screen
CLS

REM Run script
<<<<<<< HEAD
python %script_location%
=======
python %script_location%

REM Pause at the end for testing purposes
pause
>>>>>>> f184b098eed66b918e9cee3b431c4a5248cd4ec9
