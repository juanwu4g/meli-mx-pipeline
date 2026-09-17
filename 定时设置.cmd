@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM  定时任务设置：把「每月自动出报表」配进 Windows 计划任务。
REM  只需要配一次，配完这个窗口就不用再开了。
REM  日常手动跑报表请双击 启动.cmd。
REM ============================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
  echo [x] Not found: .venv\Scripts\pythonw.exe
  echo [x] 找不到虚拟环境，请先在本目录依次执行这两条命令：
  echo.
  echo     python -m venv .venv
  echo     .venv\Scripts\pip install -r requirements.txt
  echo.
  pause
  exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" schedule_gui.py
