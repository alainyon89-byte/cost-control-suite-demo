@echo off
REM ==========================================================================
REM  Cost Control Suite - prepare the offline pack      (c) 2026 Alain Yon
REM  Run this on a PC with internet where the app is already installed. It
REM  fills the "offline\wheels" folder of this package with the packages.
REM  Copy the whole package to the PC without internet, install Python 3.13
REM  there yourself, and run Install_Cost_Control_Suite.bat: it uses
REM  "offline" by itself when it cannot download.
REM ==========================================================================
setlocal
cd /d "%~dp0"
title Cost Control Suite - Prepare offline pack

set "OFFLINE=%~dp0..\offline"
set "PY=%LOCALAPPDATA%\CostControlSuiteDemo\venv\Scripts\python.exe"

if not exist "%PY%" (
    echo.
    echo  The app is not installed on this PC yet. Run Install_Cost_Control_Suite.bat
    echo  here first ^(it needs internet once^), then run this again.
    goto failed
)
if not exist "%OFFLINE%" mkdir "%OFFLINE%"
if not exist "%OFFLINE%\wheels" mkdir "%OFFLINE%\wheels"

echo.
echo  Downloading the packages for Python 3.13 and 3.14 ...
"%PY%" -m pip download -r "%~dp0requirements.txt" --only-binary=:all: --platform win_amd64 --python-version 3.13 -d "%OFFLINE%\wheels" --disable-pip-version-check
if errorlevel 1 goto failed
"%PY%" -m pip download -r "%~dp0requirements.txt" --only-binary=:all: --platform win_amd64 --python-version 3.14 -d "%OFFLINE%\wheels" --disable-pip-version-check
if errorlevel 1 echo  NOTE: the Python 3.14 set could not be completed. The pack works for Python 3.13.

echo.
echo  Done. The offline pack is in:
echo    %OFFLINE%
echo  Copy the whole package folder to the other PC and run the installer there.
echo  The Python installer itself is not in the pack: copy that across yourself.
echo.
pause
exit /b 0

:failed
echo.
echo  The offline pack could not be completed. Check the internet connection and run this again.
echo.
pause
exit /b 1
