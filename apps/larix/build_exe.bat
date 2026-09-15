@echo off
setlocal EnableExtensions
cd /d "%~dp0"
for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"

py -m pip install pyinstaller >nul
py -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name Larix_Platform_Plugin ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%PROJECT_ROOT%" ^
  Larix_User_Platform.py
if errorlevel 1 goto :error

if exist "dist\assets" rmdir /s /q "dist\assets"
xcopy /E /I /Y "%ASSETS%" "dist\assets" >nul
xcopy /E /I /Y "%PROJECT_ROOT%\shared" "dist\shared" >nul

echo.
echo Done: dist\Larix_Platform_Plugin.exe
echo Shared assets: dist\assets
pause
exit /b 0

:error
echo.
echo [ERROR] Build failed.
pause
exit /b 1
