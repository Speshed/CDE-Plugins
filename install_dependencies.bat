@echo off
setlocal EnableExtensions
cd /d "%~dp0"
where py >nul 2>&1
if not errorlevel 1 (
  py -m pip install -r "%~dp0config\requirements.txt"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] Python not found. Install Python and add it to PATH.
    goto :error
  )
  python -m pip install -r "%~dp0config\requirements.txt"
)
if errorlevel 1 goto :error
echo.
echo Dependencies installed.
pause
exit /b 0
:error
echo.
echo [ERROR] Dependency installation failed.
pause
exit /b 1
