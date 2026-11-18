@echo off
REM ==========================================================================
REM  Cost Control Suite - uninstall                      (c) 2026 Alain Yon
REM  Removes the app and its shortcuts. No admin rights needed.
REM  Your data folder (Documents\Cost Control Suite Demo) is NOT removed.
REM ==========================================================================
setlocal
cd /d "%~dp0"
title Cost Control Suite - Uninstall

set "CCS=%LOCALAPPDATA%\CostControlSuiteDemo"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo.
echo  This removes the Cost Control Suite app from this PC for you:
echo    %CCS%
echo    the Desktop and Start menu shortcuts
echo.
echo  It does NOT remove your data. Your tools, item maps, SAP data and run
echo  history stay in Documents\Cost Control Suite Demo.
echo.
set "SURE="
set /p "SURE=  Type YES to uninstall (anything else cancels): "
if /I not "%SURE%"=="YES" (
    echo.
    echo  Cancelled. Nothing was removed.
    echo.
    pause
    exit /b 0
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -Command "foreach($dir in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) { $p = Join-Path $dir 'Cost Control Suite Demo.lnk'; if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force } }"
if exist "%CCS%" rmdir /s /q "%CCS%"

echo.
if exist "%CCS%" (
    echo  Some files could not be removed. Close Cost Control Suite and run this again.
) else (
    echo  Cost Control Suite was removed.
)
echo  Your data is still in Documents\Cost Control Suite Demo. Delete that folder
echo  yourself only if you are sure you no longer need it.
echo.
pause
exit /b 0
