@echo off

REM Set repo variables
set REPO_URL=https://raw.githubusercontent.com/Gigoo25/Maintenance-Ticket-Assistant
set REPO_BRANCH=master

REM Set tool verison check variables
set CURRENT_VERSION=unidentified
set CHECK_UPDATE_VERSION=unidentified

REM Text files local variables
set Readme_Local=unidentified
set Version_Local=unidentified
set Batch_Script_Local=unidentified
set Python_Script_Local=unidentified
set Requirements_Local=unidentified
set Update_Function_Local=unidentified

REM Text files online variables
set Readme_Online=unidentified
set Version_Online=unidentified
set Batch_Script_Online=unidentified
set Python_Script_Online=unidentified
set Requirements_Online=unidentified
set Update_Function_Online=unidentified

REM Check that wget is present
if not exist "%root%\Tools\WGET\wget.exe" (
	REM Display error message that WGET was not found.
	GOTO :EOF
)

REM Check for version file
if exist "%root%\Tools\Version.txt" (
	REM Set current version as variable
	set /p CURRENT_VERSION=<%root%\Version.txt
) else (
	REM Display error message that tool version was not found.
	GOTO :EOF
)

REM Delete version check file if found
if exist "%TEMP%\Version_Check.txt" (
	del "%TEMP%\Version_Check.txt" 2>NUL
)

REM Download version to compare from online
"%root%\Tools\WGET\wget.exe" -q "%REPO_URL%/%REPO_BRANCH%/Version.txt" -O "%TEMP%\Version_Check.txt" 2>NUL
if /i %ERRORLEVEL%==0 (
	set /p CHECK_UPDATE_VERSION=<%TEMP%\Version_Check.txt
) else (
	REM Display error message that tool version was not found.
	GOTO :EOF
)

REM Check if downloaded version is greater
if "%CHECK_UPDATE_VERSION%" GTR "%CURRENT_VERSION%" (
	goto update_yes
) else (
	goto update_no
)

REM Decline update
:update_no
REM Display message that user did not accept update.
GOTO :EOF

REM Accept update
:update_yes

REM Set variables for local text files
< "%root%\Tools\Version.txt" (
	for /l %%i in (1,1,11) do set /p =
	set /p Readme_Local=
	set /p Version_Local=
	set /p Batch_Script_Local=
	set /p Python_Script_Local=
	set /p Requirements_Local=
	set /p Update_Function_Local=
)

REM Set variables for online text files
< "%TEMP%\Version_Check.txt" (
	for /l %%i in (1,1,11) do set /p =
	set /p Readme_Online=
	set /p Version_Online=
	set /p Batch_Script_Online=
	set /p Python_Script_Online=
	set /p Requirements_Online=
	set /p Update_Function_Online=
)

REM Update Readme file
if "%Readme_Online%" GTR "%Readme_Local%" (
	REM Update tool based on variables
	"%root%\Tools\WGET\wget.exe" -q  --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 "%REPO_URL%/%REPO_BRANCH%/README.md" -O "%root%\README.md"
)

REM Update Version file
if "%Version_Online%" GTR "%Version_Local%" (
	REM Update tool based on variables
	"%root%\Tools\WGET\wget.exe" -q  --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 "%REPO_URL%/%REPO_BRANCH%/Version.txt" -O "%root%\Tools\Version.txt"
)

REM Update Batch Script
if "%Batch_Script_Online%" GTR "%Batch_Script_Local%" (
	REM Update tool based on variables
	REM "%root%\Tools\WGET\wget.exe" --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 --progress=bar:force "%REPO_URL%/%REPO_BRANCH%/Requirements.txt" -O "%root%\Tools\Requirements.txt" 2>NUL
)

REM Update Python Script
if "%Python_Script_Online%" GTR "%Python_Script_Local%" (
	REM Update tool based on variables
	REM "%root%\Tools\WGET\wget.exe" --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 --progress=bar:force "%REPO_URL%/%REPO_BRANCH%/script.py" -O "%root%\script.py" 2>NUL
)

REM Update Requirements
if "%Requirements_Online%" GTR "%Requirements_Local%" (
	REM Update tool based on variables
	REM "%root%\Tools\WGET\wget.exe" --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 --progress=bar:force "%REPO_URL%/%REPO_BRANCH%/requirements.txt" -O "%root%\requirements.txt" 2>NUL
)

REM Update "Update Function"
if "%Update_Function_Online%" GTR "%Update_Function_Local%" (
	REM Update tool based on variables
	REM "%root%\Tools\WGET\wget.exe" --retry-connrefused --waitretry=1 --read-timeout=20 --timeout=15 -t 2 --progress=bar:force "%REPO_URL%/%REPO_BRANCH%/Tools/Functions/Update_Function.bat" -O "%root%\Tools\Functions\Update_Function.bat" 2>NUL
)

REM Display message that update was successful.
pause
exit /b

