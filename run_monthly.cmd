@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM  每月 15 号由计划任务调用：下载上个月数据 + 生成上个月报表。
REM  手动双击也可以跑。日志在 logs\monthly_<YYYYMM>.log。
REM
REM  退出码：0 成功 / 1 有需关注项 / 2 紫鸟错误 / 3 白名单
REM          4 配置错误 / 5 上一次还在跑
REM ============================================================
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [x] Not found: .venv\Scripts\python.exe
  echo [x] 缺少虚拟环境，请先在本目录依次执行这两条命令：
  echo     python -m venv .venv
  echo     .venv\Scripts\pip install -r requirements.txt
  exit /b 4
)

REM -u 关掉输出缓冲：计划任务超时会强杀进程，缓冲区里的日志会全丢。
".venv\Scripts\python.exe" -u run_monthly.py %*
exit /b %ERRORLEVEL%
