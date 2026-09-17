@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Run run.bat first.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 (
    echo Failed to install PyInstaller.
    pause
    exit /b 1
)

rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

".venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name ProjectPointTest ^
  --collect-all PySide6 ^
  main.py

if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done.
echo EXE folder: %CD%\dist\ProjectPointTest
pause
endlocal
