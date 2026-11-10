@echo off
setlocal
cd /d "%~dp0"
set "NSE_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%NSE_PYTHON%" (
  "%NSE_PYTHON%" update_weekly_ohlc.py --auto-holidays %*
) else (
  py -3 update_weekly_ohlc.py --auto-holidays %*
)
if errorlevel 1 echo Update failed. See the error above. Existing completed workbooks were preserved.
pause
