@echo off
setlocal EnableExtensions
cd /d "%~dp0"

for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"
set "APP_NAME=VitroCAD"
set "ENTRY_FILE=VitroCAD.py"
set "DIST_DIR=%~dp0dist"
set "BUILD_DIR=%~dp0build"
title Build VitroCAD Tools

set "HTTP_PROXY="
set "HTTPS_PROXY="
set "ALL_PROXY="
set "NO_PROXY=*"
set "http_proxy="
set "https_proxy="
set "all_proxy="
set "no_proxy=*"

echo ============================================
echo   Building %APP_NAME%.exe
echo ============================================
echo.

if not exist "%ENTRY_FILE%" (
    echo [ERROR] Entry file not found: %ENTRY_FILE%
    pause
    exit /b 1
)
if not exist "%ASSETS%\icon.ico" (
    echo [ERROR] Shared assets folder not found: %ASSETS%
    pause
    exit /b 1
)

where pyinstaller >nul 2>&1
if not errorlevel 1 (
    set "PYINSTALLER_CMD=pyinstaller"
    goto :build
)
py -m PyInstaller --version >nul 2>&1
if not errorlevel 1 (
    set "PYINSTALLER_CMD=py -m PyInstaller"
    goto :build
)
python -m PyInstaller --version >nul 2>&1
if not errorlevel 1 (
    set "PYINSTALLER_CMD=python -m PyInstaller"
    goto :build
)

echo [ERROR] PyInstaller not found.
echo Run: py -m pip install -r requirements.txt
pause
exit /b 1

:build
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"

%PYINSTALLER_CMD% ^
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
    --collect-all PySide6 ^
    --hidden-import requests ^
    --hidden-import urllib3 ^
    --hidden-import urllib3.contrib.socks ^
    --hidden-import socks ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    "%ENTRY_FILE%"

if errorlevel 1 goto :error
xcopy /E /I /Y "%ASSETS%" "%DIST_DIR%\assets" >nul
xcopy /E /I /Y "%PROJECT_ROOT%\shared" "%DIST_DIR%\shared" >nul

echo.
echo Done: dist\%APP_NAME%.exe
echo Shared assets: dist\assets
pause
exit /b 0

:error
echo.
echo [ERROR] Build failed.
pause
exit /b 1
