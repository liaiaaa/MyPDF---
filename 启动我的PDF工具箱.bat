@echo off
setlocal
title MySafePDFToolbox Launcher
cd /d "%~dp0"

rem ============================================================
rem  Launcher: pin to the Python that has all deps installed.
rem  (File is GBK-encoded to match Chinese Windows console.)
rem ============================================================

set "PYW="

rem 1) Preferred: dedicated Python with all libs
if exist "E:\software\python\pythonw.exe" set "PYW=E:\software\python\pythonw.exe"
if not defined PYW if exist "E:\software\python\python.exe" set "PYW=E:\software\python\python.exe"

rem 2) Fallbacks
if not defined PYW where pyw >nul 2>nul && set "PYW=pyw"
if not defined PYW where pythonw >nul 2>nul && set "PYW=pythonw"

if not defined PYW (
    echo [错误] 没有找到可用的 Python 环境。
    echo 请安装 Python 3.9 或以上版本，然后重新双击本脚本。
    echo.
    pause
    exit /b 1
)

echo 正在启动「我的专属安全 PDF 工具箱」...
"%PYW%" pdf_toolbox.py > "run_log.txt" 2>&1
if errorlevel 1 (
    echo.
    echo [错误] 程序启动失败，详情已写入：run_log.txt
    echo 请把该文件内容发给我排查。
    echo.
    pause
)
endlocal
