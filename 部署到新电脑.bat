@echo off
setlocal
title Deploy MyPDFToolbox
cd /d "%~dp0"

echo ============================================
echo   我的专属安全 PDF 工具箱 - 一键部署
echo   （在新电脑上运行一次即可）
echo ============================================
echo.

rem ---- 1. 查找 Python ----
set "PY="
where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
    echo.
    echo [错误] 未找到 Python 环境。
    echo 请先安装 Python 3.9 或以上版本：
    echo   官网下载：https://www.python.org/downloads/
    echo   安装时务必勾选 "Add Python to PATH"
    echo 安装完成后，重新双击本脚本即可。
    echo.
    pause
    exit /b 1
)

echo [1/3] 检查依赖库是否已安装...
"%PY%" -c "import fitz,PIL,docx,openpyxl,pptx,pdfplumber" >nul 2>nul
if errorlevel 1 (
    echo [2/3] 正在安装依赖（需要联网，首次约 1-2 分钟）...
    "%PY%" -m pip install -r "requirements.txt"
    if errorlevel 1 (
        echo.
        echo [错误] 依赖安装失败，请检查网络后重试。
        pause
        exit /b 1
    )
) else (
    echo       依赖已就绪，跳过安装。
)

rem ---- 3. 创建桌面快捷方式（由 Python 辅助完成，兼容中文路径）----
echo [3/3] 创建桌面快捷方式...
"%PY%" make_shortcut.py >nul 2>nul
if errorlevel 1 (
    echo       快捷方式创建失败（不影响使用，可改双击「启动我的PDF工具箱.bat」）。
) else (
    echo       桌面快捷方式已创建：「我的PDF工具箱」
)

echo.
echo ============ 部署完成 ============
echo 现在可以双击桌面「我的PDF工具箱」使用了。
echo 也可以直接把 PDF / 图片拖到窗口或图标上快速处理。
echo.
pause
endlocal
