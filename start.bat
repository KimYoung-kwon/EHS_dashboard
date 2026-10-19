@echo off
rem AX-FA dashboard server launcher (double-click to run)
chcp 65001 >nul
cd /d "%~dp0"
where py >/dev/null 2>nul
if %errorlevel%==0 (
  py -3 serve.py %*
  goto done
)
where python >/dev/null 2>nul
if %errorlevel%==0 (
  python serve.py %*
  goto done
)
echo Python 3 is not installed. Install it from https://www.python.org/downloads/ and check "Add python.exe to PATH".
:done
pause
