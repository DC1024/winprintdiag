@echo off
chcp 936 >nul 2>&1
cd /d "%~dp0"
setlocal
set "THISDIR=%~dp0"

if not exist "%THISDIR%WinPrintDiag.exe" goto ps1

"%THISDIR%WinPrintDiag.exe" -OpenReport %*
if errorlevel 1 (
    echo.
    echo   命令行版（exe）启动失败，自动改用 PowerShell 脚本方式执行……
    echo.
    goto ps1
)
goto msg

:ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%THISDIR%WinPrintDiag.ps1" -OpenReport %*

:msg
echo.
echo ------------------------------------------------------------
echo  体检完成，报告已自动打开并保存在本目录下。
echo  需要修复？右键本文件，选择"以管理员身份运行"，
echo  并在参数末尾追加 -Repair（清空打印队列改用 -ClearQueue）。
echo ------------------------------------------------------------
pause
endlocal
