@echo off
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>&1 && (
  py -u "数据分析.py"
  goto :done
)
where python >nul 2>&1 && (
  python -u "数据分析.py"
  goto :done
)

echo 未找到 Python，请先安装并勾选 "Add to PATH"，或使用 py 启动器。
pause
exit /b 1

:done
exit /b %errorlevel%
