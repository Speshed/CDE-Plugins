@echo off
setlocal EnableExtensions
cd /d "%~dp0"
for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"

where py >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=py"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found.
        goto :error
    )
    set "PYTHON=python"
)

%PYTHON% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller is not installed. Run install_dependencies.bat first.
    goto :error
)

if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

%PYTHON% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "SGNL_Platform" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%PROJECT_ROOT%" ^
  SIGNAL.py
if errorlevel 1 goto :error

if not exist "dist\SGNL_Platform\SGNL_Platform.exe" (
    echo [ERROR] Expected output not found: dist\SGNL_Platform\SGNL_Platform.exe
    goto :error
)

echo.
echo Done: dist\SGNL_Platform\SGNL_Platform.exe
if not defined CDE_BUILD_ALL pause
exit /b 0

:error
echo.
echo [ERROR] SIGNAL Docs build failed.
if not defined CDE_BUILD_ALL pause
exit /b 1
