@echo off
chcp 936 >nul 2>&1
cd /d "%~dp0"
setlocal
set "THISDIR=%~dp0"

if not exist "%THISDIR%WinPrintDiagUI.exe" goto ps1

"%THISDIR%WinPrintDiagUI.exe" %*
if errorlevel 1 (
    echo.
    echo   图形界面（exe）启动失败，自动改用 PowerShell 脚本方式启动……
    echo.
    goto ps1
)
goto end

:ps1
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%THISDIR%WinPrintDiagUI.ps1" %*

:end
endlocal
