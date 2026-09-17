@echo off
setlocal EnableExtensions
cd /d "%~dp0"
for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"
set "APP_NAME=VitroCAD"
set "ENTRY_FILE=VitroCAD.py"
set "DIST_DIR=%~dp0dist"
set "BUILD_DIR=%~dp0build"

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

if not exist "%ENTRY_FILE%" (
    echo [ERROR] Entry file not found: %ENTRY_FILE%
    goto :error
)
if not exist "%ASSETS%\icon.ico" (
    echo [ERROR] Shared assets folder not found: %ASSETS%
    goto :error
)

if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"

%PYTHON% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "%APP_NAME%" ^
    --icon "%ASSETS%\icon.ico" ^
    --paths "%PROJECT_ROOT%" ^
    --distpath "%DIST_DIR%" ^
    --workpath "%BUILD_DIR%" ^
    --specpath "%BUILD_DIR%" ^
    --hidden-import requests ^
    --hidden-import urllib3 ^
    --hidden-import urllib3.contrib.socks ^
    --hidden-import socks ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    "%ENTRY_FILE%"
if errorlevel 1 goto :error

if not exist "dist\VitroCAD.exe" (
    echo [ERROR] Expected output not found: dist\VitroCAD.exe
    goto :error
)

echo.
echo Done: dist\VitroCAD.exe
if not defined CDE_BUILD_ALL pause
exit /b 0

:error
echo.
echo [ERROR] Vitro build failed.
if not defined CDE_BUILD_ALL pause
exit /b 1
