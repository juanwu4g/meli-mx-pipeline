@echo off
chcp 65001 >nul 2>&1
REM ---------------------------------------------------------------
REM 双击这个文件打开操作界面。
REM
REM chcp 65001 是把控制台切到 UTF-8：这个文件用 UTF-8 存，而 cmd.exe
REM 默认用 OEM 代码页（中文 Windows 是 936），不切的话下面的提示会显示
REM 成乱码 —— 偏偏那是 .venv 缺失时用户唯一能看到的指引。
REM
REM 解释器写死 .venv 里的：机器上常有别的 Python（conda 之类），
REM requirements 锁的是 pandas 3，用成 pandas 2 会让报表崩在库存页。
REM ---------------------------------------------------------------
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
  echo.
  echo   [x] Not found: .venv\Scripts\pythonw.exe
  echo   [x] 找不到虚拟环境，请先在本目录依次执行这两条命令：
  echo.
  echo       python -m venv .venv
  echo       .venv\Scripts\pip install -r requirements.txt
  echo.
  pause
  exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" gui.py
