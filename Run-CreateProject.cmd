@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_EXE=pwsh"
where pwsh.exe >nul 2>nul
if errorlevel 1 set "PS_EXE=powershell"

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%CreateProject.ps1"

echo.
pause
