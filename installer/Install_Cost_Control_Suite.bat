@echo off
REM ==========================================================================
REM  Cost Control Suite - installer / updater            (c) 2026 Alain Yon
REM  No admin rights needed. Run it again to update: your data is kept.
REM
REM  Python is NOT installed by this file. Install Python 3.13 (64-bit) from
REM  python.org yourself first ("Install just for me" is fine). This installer
REM  finds it in the usual folders, or asks you for the path, and builds the
REM  app its own package folder on top of it, so nothing is added to your
REM  Python itself.
REM
REM  Rules kept on purpose: no delayed expansion, every path quoted, Python is
REM  always called by its full path (never searched for), plain ASCII, CRLF.
REM ==========================================================================
setlocal
cd /d "%~dp0"
title Cost Control Suite - Install

set "PKG=%~dp0.."
set "CCS=%LOCALAPPDATA%\CostControlSuiteDemo"
set "VENV=%CCS%\venv"
set "PY=%VENV%\Scripts\python.exe"
set "PYW=%VENV%\Scripts\pythonw.exe"
set "LOG=%CCS%\logs\install_log.txt"
set "OFFLINE=%PKG%\offline"
set "CURL=%SystemRoot%\System32\curl.exe"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "REG=%SystemRoot%\System32\reg.exe"
set "VERCHECK=import sys; sys.exit(0 if sys.version_info[:2] in ((3, 13), (3, 14)) and sys.maxsize > 2**32 else 1)"
set "FAILWHY="

for /F %%a in ('echo prompt $E ^| "%ComSpec%"') do set "ESC=%%a"
set "C_HEAD=%ESC%[97;44m"
set "C_OK=%ESC%[92m"
set "C_WARN=%ESC%[93m"
set "C_BAD=%ESC%[91m"
set "C_DIM=%ESC%[90m"
set "C_OFF=%ESC%[0m"

echo.
echo %C_HEAD%                                                                      %C_OFF%
echo %C_HEAD%   COST CONTROL SUITE  -  Install                                     %C_OFF%
echo %C_HEAD%   Demo Hotel - generated sample data         (c) 2026 Alain Yon      %C_OFF%
echo %C_HEAD%                                                                      %C_OFF%
echo.
echo   This sets up the app for you only. No admin rights are needed.
echo   Run it again any time to update. Your data folder is never replaced.
echo.

if not exist "%CCS%" mkdir "%CCS%"
if not exist "%CCS%\logs" mkdir "%CCS%\logs"
echo ==== Install started %DATE% %TIME% ==== >> "%LOG%"

REM ------------------------------------------------------------------ step 1
echo %C_HEAD% 1 of 6 %C_OFF%  Check this PC
if /I not "%PROCESSOR_ARCHITECTURE%"=="AMD64" if /I not "%PROCESSOR_ARCHITEW6432%"=="AMD64" (
    set "FAILWHY=This PC is not running 64-bit Windows. The app needs 64-bit Windows 10 or 11."
    goto failed
)
echo          %C_OK%OK%C_OFF%    64-bit Windows
echo test> "%CCS%\write_test.tmp"
if not exist "%CCS%\write_test.tmp" (
    set "FAILWHY=Cannot write to %CCS%. Ask IT why your own AppData folder is read-only."
    goto failed
)
del "%CCS%\write_test.tmp" >nul 2>&1
echo          %C_OK%OK%C_OFF%    Can write to your app folder
set "NET=yes"
"%CURL%" -s -I -m 15 -o nul "https://pypi.org" >nul 2>&1
if errorlevel 1 set "NET=no"
if "%NET%"=="yes" echo          %C_OK%OK%C_OFF%    Internet reachable
if "%NET%"=="no" echo          %C_WARN%NOTE%C_OFF%  No internet. The offline pack is used if it is in this package.
set "WV2=no"
"%REG%" query "HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" /v pv >nul 2>&1 && set "WV2=yes"
"%REG%" query "HKCU\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}" /v pv >nul 2>&1 && set "WV2=yes"
if "%WV2%"=="yes" echo          %C_OK%OK%C_OFF%    WebView2 is installed
if "%WV2%"=="no" echo          %C_WARN%NOTE%C_OFF%  WebView2 not found. The app will open in an Edge window instead.
echo step 1 ok, internet=%NET% webview2=%WV2% >> "%LOG%"
echo.

REM ------------------------------------------------------------------ step 2
echo %C_HEAD% 2 of 6 %C_OFF%  Python ^(the one you installed; nothing is added to it^)
REM An app folder that already works is kept as it is.
if exist "%PY%" "%PY%" -c "%VERCHECK%" >nul 2>&1 && goto havepython

set "BASEPY="
if exist "%CCS%\python\python.exe" set "BASEPY=%CCS%\python\python.exe"
if not defined BASEPY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "BASEPY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined BASEPY if exist "%LOCALAPPDATA%\Python\pythoncore-3.13-64\python.exe" set "BASEPY=%LOCALAPPDATA%\Python\pythoncore-3.13-64\python.exe"
if not defined BASEPY if exist "%ProgramFiles%\Python313\python.exe" set "BASEPY=%ProgramFiles%\Python313\python.exe"
if not defined BASEPY if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" set "BASEPY=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
if not defined BASEPY if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe" set "BASEPY=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
if not defined BASEPY if exist "%ProgramFiles%\Python314\python.exe" set "BASEPY=%ProgramFiles%\Python314\python.exe"
if defined BASEPY goto checkbase

echo          %C_WARN%NOTE%C_OFF%  Python 3.13 or 3.14 was not found in the usual folders.
echo                If it is not installed yet: get "Windows installer (64-bit)" for
echo                Python 3.13 from python.org, install it, then run this again.
echo                If it is installed somewhere else, type the full path to its
echo                python.exe below ^(or just press Enter to stop^).
echo.
set /p "BASEPY=  Path to python.exe: "
if not defined BASEPY (
    set "FAILWHY=Python is not installed yet. Install Python 3.13 (64-bit) from python.org, then run this installer again."
    goto failed
)
set "BASEPY=%BASEPY:"=%"
if not exist "%BASEPY%" (
    set "FAILWHY=That python.exe was not found. Check the path and run this again."
    goto failed
)

:checkbase
"%BASEPY%" -c "%VERCHECK%" >nul 2>&1
if errorlevel 1 (
    set "FAILWHY=The Python at %BASEPY% is not a 64-bit Python 3.13 or 3.14. The tools are tested on 3.13 (3.14 also works). Install Python 3.13 (64-bit) from python.org and run this again."
    goto failed
)
echo          Using %BASEPY%
echo base python %BASEPY% >> "%LOG%"
if exist "%VENV%" rmdir /s /q "%VENV%"
"%BASEPY%" -m venv "%VENV%" >> "%LOG%" 2>&1
if not exist "%PY%" (
    set "FAILWHY=The app's own package folder could not be created from %BASEPY%. See the install log."
    goto failed
)

:havepython
"%PY%" --version >> "%LOG%" 2>&1
if errorlevel 1 (
    set "FAILWHY=The app's Python in %VENV% does not start. Delete that folder and run this again."
    goto failed
)
echo          %C_OK%OK%C_OFF%    Python is ready
echo.

REM ------------------------------------------------------------------ step 3
echo %C_HEAD% 3 of 6 %C_OFF%  Packages ^(a few minutes the first time^)
"%PY%" -c "import sys, importlib.metadata as m; want=[l.split('==') for l in open(sys.argv[1]).read().split() if '==' in l]; sys.exit(0 if all(m.version(n)==v for n,v in want) else 1)" "%~dp0requirements.txt" >nul 2>&1
if not errorlevel 1 goto packagesok
if "%NET%"=="no" goto pipoffline
"%PY%" -m pip install --upgrade pip --disable-pip-version-check --no-warn-script-location >> "%LOG%" 2>&1
"%PY%" -m pip install -r "%~dp0requirements.txt" --disable-pip-version-check --no-warn-script-location >> "%LOG%" 2>&1
if not errorlevel 1 goto packagesok
echo          %C_WARN%NOTE%C_OFF%  pip could not connect directly. Trying the Windows proxy.
set "PROXY="
for /F "usebackq delims=" %%p in (`"%PS%" -NoProfile -Command "$u=[Uri]'https://pypi.org'; $p=[System.Net.WebRequest]::GetSystemWebProxy().GetProxy($u); if ($p.Host -ne $u.Host) { $p.AbsoluteUri }"`) do set "PROXY=%%p"
if not defined PROXY goto pipoffline
echo proxy %PROXY% >> "%LOG%"
"%PY%" -m pip install -r "%~dp0requirements.txt" --proxy "%PROXY%" --disable-pip-version-check --no-warn-script-location >> "%LOG%" 2>&1
if not errorlevel 1 goto packagesok

:pipoffline
if not exist "%OFFLINE%\wheels" (
    set "FAILWHY=The packages could not be downloaded and there is no offline pack in this package. Run Prepare_Offline_Pack.bat on a PC with internet, copy the offline folder into this package, and run this again."
    goto failed
)
echo          Using the offline pack ...
"%PY%" -m pip install --no-index --find-links "%OFFLINE%\wheels" -r "%~dp0requirements.txt" --disable-pip-version-check --no-warn-script-location >> "%LOG%" 2>&1
if errorlevel 1 (
    set "FAILWHY=The packages could not be installed, even from the offline pack."
    goto failed
)

:packagesok
echo          %C_OK%OK%C_OFF%    Packages are installed
echo.

REM ------------------------------------------------------------------ step 4
echo %C_HEAD% 4 of 6 %C_OFF%  Copy files
"%PY%" "%~dp0install_steps.py" copy "%PKG%"
if errorlevel 1 (
    set "FAILWHY=The files could not be copied. See the lines above."
    goto failed
)
echo step 4 ok >> "%LOG%"
echo          %C_OK%OK%C_OFF%    Files are in place. Existing data files were not replaced.
echo.

REM ------------------------------------------------------------------ step 5
echo %C_HEAD% 5 of 6 %C_OFF%  Shortcuts
"%PS%" -NoProfile -ExecutionPolicy Bypass -Command "$ccs=$env:CCS; $w=New-Object -ComObject WScript.Shell; foreach($dir in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) { $s=$w.CreateShortcut((Join-Path $dir 'Cost Control Suite Demo.lnk')); $s.TargetPath=(Join-Path $ccs 'venv\Scripts\pythonw.exe'); $s.Arguments='\"' + (Join-Path $ccs 'app\main.py') + '\"'; $s.WorkingDirectory=$ccs; $s.IconLocation=(Join-Path $ccs 'app\web\icon.ico'); $s.Description='Cost Control Suite'; $s.Save() }" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo          %C_WARN%NOTE%C_OFF%  Shortcuts could not be made. Start the app with:
    echo                "%PYW%" "%CCS%\app\main.py"
) else (
    echo          %C_OK%OK%C_OFF%    Desktop and Start menu: Cost Control Suite Demo
)
echo.

REM ------------------------------------------------------------------ step 6
echo %C_HEAD% 6 of 6 %C_OFF%  Self-test
"%PY%" "%CCS%\app\main.py" --selftest > "%CCS%\logs\selftest.txt" 2>&1
set "ST=%ERRORLEVEL%"
type "%CCS%\logs\selftest.txt"
type "%CCS%\logs\selftest.txt" >> "%LOG%"
if not "%ST%"=="0" (
    set "FAILWHY=The self-test failed. See the FAIL lines above."
    goto failed
)
if not exist "%PKG%\source_tools" goto done
echo.
echo          Running each tool once on its sample files ^(this takes a few minutes^) ...
"%PY%" "%PKG%\tests\run_regression.py" --quick --source "%PKG%\source_tools" > "%CCS%\logs\quick_regression.txt" 2>&1
set "RG=%ERRORLEVEL%"
"%SystemRoot%\System32\findstr.exe" /B /C:"PASS" /C:"FAIL" /C:"SKIP" /C:"NOT TESTED" /C:"ALL PASSED" "%CCS%\logs\quick_regression.txt"
type "%CCS%\logs\quick_regression.txt" >> "%LOG%"
if not "%RG%"=="0" (
    set "FAILWHY=One or more tools did not pass on their sample files. The app is installed, but send Alain the install log before relying on those tools."
    goto failed
)

:done
echo ==== Install finished OK %DATE% %TIME% ==== >> "%LOG%"
echo.
echo %C_OK%  Done. Open Cost Control Suite Demo from your Desktop.%C_OFF%
echo %C_DIM%  Your data: Documents\Cost Control Suite Demo      Install log: %LOG%%C_OFF%
echo.
pause
exit /b 0

:failed
echo ==== Install FAILED: %FAILWHY% ==== >> "%LOG%"
echo.
echo %C_BAD%  The install did not finish.%C_OFF%
echo   %FAILWHY%
echo.
echo   What to send Alain: the file
echo     %LOG%
echo.
pause
exit /b 1
