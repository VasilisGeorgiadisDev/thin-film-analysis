@echo off
chcp 65001 <nul >nul
set PYTHONUTF8=1
cd /d "%~dp0"
set "PY="
py -3 --version <nul >nul 2>&1 && set "PY=py -3"
if not defined PY (python --version <nul >nul 2>&1 && set "PY=python")
if not defined PY (
    echo.
    echo   Python was not found on this computer.
    echo.
    echo   1. Download it from https://www.python.org/downloads/
    echo   2. During setup, tick "Add python.exe to PATH".
    echo   3. Double-click this file again.
    echo.
    pause
    exit /b 1
)
%PY% "%~dp0menu.py" %*
echo.
pause
