@echo off
chcp 65001 >nul
title BOM 成本分析 - 网页版
cd /d "%~dp0"

echo ========================================
echo   BOM 成本分析 - 网页版（本目录独立运行）
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto try_py
set "PY=python"
goto have_py
:try_py
py --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python。
    echo 请安装: https://www.python.org/downloads/
    pause
    exit /b 1
)
set "PY=py"
:have_py

echo [1/2] 安装/检查依赖 ^(本目录 requirements.txt^)...
"%PY%" -m pip install -r "%~dp0requirements.txt" -q
if errorlevel 1 (
    echo [错误] pip 安装失败。
    pause
    exit /b 1
)

echo [2/2] 启动服务...
echo.
echo   启动后将自动打开浏览器 ^(约 1 秒^)
echo   手动访问: http://127.0.0.1:5000
echo   停止服务: 在本窗口按 Ctrl+C
echo.
echo ========================================
echo.

"%PY%" "%~dp0app.py"
echo.
echo 服务已结束。
pause
