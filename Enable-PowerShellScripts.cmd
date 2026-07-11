@echo off
setlocal

echo Enabling local PowerShell scripts for the current user...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force; Write-Host 'Windows PowerShell ExecutionPolicy CurrentUser=RemoteSigned is set.' -ForegroundColor Green"
where pwsh.exe >nul 2>nul
if not errorlevel 1 (
    pwsh -NoProfile -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force; Write-Host 'PowerShell 7 ExecutionPolicy CurrentUser=RemoteSigned is set.' -ForegroundColor Green"
)

echo.
echo You can now run:
echo   .\CreateProject.ps1
echo.
pause
