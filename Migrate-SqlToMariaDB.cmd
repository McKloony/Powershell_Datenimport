@echo off
setlocal EnableExtensions

REM ---------------------------------------------------------------------------
REM  Migrate-SqlToMariaDB.cmd
REM
REM  Host launcher for Migrate-SqlToMariaDB.ps1.
REM
REM  Resolves the PowerShell script next to this .cmd file (so the launcher can
REM  be double-clicked or pinned) and forwards every argument 1:1 to it.
REM
REM  Usage:
REM    Migrate-SqlToMariaDB.cmd
REM        -> opens the interactive menu
REM    Migrate-SqlToMariaDB.cmd -Mode Verify -SkipBackup
REM    Migrate-SqlToMariaDB.cmd -Mode Migrate -SkipBackup
REM    Migrate-SqlToMariaDB.cmd -Mode Repair
REM    Migrate-SqlToMariaDB.cmd -Mode Cleanup -Force
REM ---------------------------------------------------------------------------

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%Migrate-SqlToMariaDB.ps1"

if not exist "%PS1%" (
    echo [ERROR] Cannot find PowerShell script:
    echo         %PS1%
    echo Place this .cmd file next to Migrate-SqlToMariaDB.ps1 and try again.
    exit /b 2
)

REM Prefer Windows PowerShell 5.1 (powershell.exe) because the script targets
REM that runtime (#Requires -Version 5.1) and SimplySql 2.x is shipped for it.
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell.exe"

REM Run from the script's directory so any relative paths inside the .ps1
REM resolve against the project folder, not the caller's CWD.
pushd "%SCRIPT_DIR%" >nul

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -NoLogo -File "%PS1%" %*
set "RC=%ERRORLEVEL%"

popd >nul

REM Pause when launched by double-click so the user can read the summary.
REM cmd.exe started by Explorer has CMDCMDLINE matching /c "...".
echo %CMDCMDLINE% | findstr /I /C:"/c" >nul
if not errorlevel 1 (
    echo.
    pause
)

exit /b %RC%
