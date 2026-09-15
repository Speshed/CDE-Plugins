@echo off
setlocal EnableExtensions
cd /d "%~dp0"
for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"
title Build ProjectPoint.exe

echo ============================================
echo   Building ProjectPoint.exe
echo ============================================
echo.

where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller not found.
    echo Install: pip install -r requirements.txt
    pause
    exit /b 1
)
if not exist "%ASSETS%\icon.ico" (
    echo [ERROR] Shared assets folder not found: %ASSETS%
    pause
    exit /b 1
)

if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "ProjectPoint" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%PROJECT_ROOT%" ^
  --distpath "%~dp0dist" ^
  --workpath "%~dp0build" ^
  --specpath "%~dp0build" ^
  --add-data "%~dp0projectpoint\resources\builtin_role_permissions.json;projectpoint\resources" ^
  "%~dp0main.py"

if errorlevel 1 goto :error
if exist "%~dp0connection_profiles.json" copy /Y "%~dp0connection_profiles.json" "%~dp0dist\connection_profiles.json" >nul
xcopy /E /I /Y "%ASSETS%" "%~dp0dist\assets" >nul
xcopy /E /I /Y "%PROJECT_ROOT%\shared" "%~dp0dist\shared" >nul

echo.
echo Build complete: dist\ProjectPoint.exe
echo Profiles:       dist\connection_profiles.json
echo Shared images:  dist\assets
pause
exit /b 0

:error
echo.
echo [ERROR] Build failed.
pause
exit /b 1
