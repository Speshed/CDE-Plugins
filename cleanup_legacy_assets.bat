@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist "launcher.py" goto :wrong_root
if not exist "assets\icon.ico" goto :missing_assets

echo Removing legacy asset folders...
call :remove_dir "launcher_assets"
call :remove_dir "apps\larix\larix_assets"
call :remove_dir "apps\larix\coll\icon"
call :remove_dir "apps\signal\assets"
call :remove_dir "apps\vitrocad\assets"
call :remove_dir "apps\projectpoint\projectpoint\resources\ui"

echo.
echo Done. Shared external images are now stored only in the root assets folder.
pause
exit /b 0

:remove_dir
if exist "%~1" rmdir /s /q "%~1"
exit /b 0

:wrong_root
echo [ERROR] Run this file from the SOD_Manager project root.
pause
exit /b 1

:missing_assets
echo [ERROR] Shared assets folder is missing or incomplete.
echo Expected file: assets\icon.ico
pause
exit /b 1
