@echo off
:: Enable ANSI escape sequences (Windows 10 or later)
:: Older consoles might not support this coloring.

:: Define color codes
set "RED=\x1b[31m"
set "YELLOW=\x1b[33m"
set "RESET=\x1b[0m"

:: Print header in RED and YELLOW
echo %RED%=============================================%RESET%
echo %YELLOW%Starting the Download Helper using PowerShell...%RESET%
echo %RED%=============================================%RESET%

echo.

:: Check legacy Windows PowerShell version
echo Checking Windows PowerShell version...
powershell -NoProfile -Command "$PSVersionTable"

echo.

:: Check if pwsh (PowerShell Core) exists
where pwsh >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo pwsh.exe found!
    echo.
    pwsh -NoProfile -Command "$PSVersionTable"
) else (
    echo pwsh.exe not found on this system.
)

echo.
echo Running the helper script...

:: Run the script using Windows PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Start-DownloadHelper.ps1"
