@echo off
setlocal EnableExtensions
cd /d "%~dp0"
where py >nul 2>&1
if not errorlevel 1 (
  set "PYTHON=py"
) else (
  where python >nul 2>&1
  if errorlevel 1 goto :no_python
  set "PYTHON=python"
)
%PYTHON% app\launcher.py
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%

:no_python
echo [ERROR] Python not found. Install Python and add it to PATH.
pause
exit /b 1
