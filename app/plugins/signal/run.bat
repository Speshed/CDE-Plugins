@echo off
setlocal
cd /d "%~dp0"
python SIGNAL.py
if errorlevel 1 (
  echo.
  echo Application exited with an error.
  pause
)
