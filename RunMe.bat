@echo off

REM Minimize Window
REM if not "%1" == "min" start /MIN cmd /c %0 min & exit/b

REM Set location of directory
%~d0
cd %~dp0
set ROOT=%cd%
set Maintenance_Assistant_Location=%ROOT%\Maintenance_Assistant.py
set Update_Script_Location=%ROOT%\Tools\Functions\Update_Script.py

REM Install prerequisites
pip install -r %ROOT%\Tools\Requirements.txt

REM Clear screen
CLS

REM Run update
python %Update_Script_Location%
pause
REM Clear screen
CLS

REM Run assistant
python %Maintenance_Assistant_Location%