@echo off
setlocal

set "MODE=%~1"
set "PS_EXE=pwsh"
where pwsh.exe >nul 2>nul
if errorlevel 1 set "PS_EXE=powershell"

if /I "%MODE%"=="--admin-only" goto AdminOnly

echo AsiToPix initialization
echo.

echo [1/2] Enabling local PowerShell scripts for the current user...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force; Write-Host 'Windows PowerShell ExecutionPolicy CurrentUser=RemoteSigned is set.' -ForegroundColor Green"
where pwsh.exe >nul 2>nul
if not errorlevel 1 (
    pwsh -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force; Write-Host 'PowerShell 7 ExecutionPolicy CurrentUser=RemoteSigned is set.' -ForegroundColor Green"
)

echo [2/2] Checking Administrator rights for symlink settings...
net session >nul 2>nul
if not errorlevel 1 goto AdminSettings

echo [!] This window is not elevated.
echo     Network symlink support requires Administrator rights.
choice /C YN /N /M "Open an elevated window to enable symlink settings? [Y/N] "
if errorlevel 2 goto Done

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath 'cmd.exe' -ArgumentList '/c ""%~f0"" --admin-only' -Verb RunAs"
goto Done

:AdminOnly
echo AsiToPix elevated initialization
echo.

:AdminSettings
echo Enabling symlink evaluation...
fsutil behavior set SymlinkEvaluation L2L:1
fsutil behavior set SymlinkEvaluation L2R:1
fsutil behavior set SymlinkEvaluation R2L:1
fsutil behavior set SymlinkEvaluation R2R:1
fsutil behavior query SymlinkEvaluation

echo.
echo Enabling long path support...
reg add HKLM\SYSTEM\CurrentControlSet\Control\FileSystem /v LongPathsEnabled /t REG_DWORD /d 1 /f

echo.
echo [OK] Administrator settings applied.

:Done
echo.
echo Initialization finished.
echo.
echo You can now run:
echo   .\CreateProject.ps1
echo or:
echo   Run-CreateProject.cmd
echo.
pause
