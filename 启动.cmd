@echo off
REM 双击这个文件打开操作界面。
REM 写死 .venv 里的解释器 —— 机器上常有别的 Python（conda 之类），
REM 版本不同会让报表崩在 pandas 的 NaN 处理上。
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo 找不到 .venv\Scripts\python.exe
  echo.
  echo 先在本目录执行：
  echo     python -m venv .venv
  echo     .venv\Scripts\pip install -r requirements.txt
  echo.
  pause
  exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" gui.py
