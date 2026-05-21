@echo off
REM 网页版独立包在 web\ 下；本脚本从项目根目录转发启动。
if not exist "%~dp0web\app.py" (
    echo [错误] 未找到 web\app.py
    pause
    exit /b 1
)
call "%~dp0web\start_web.bat"
