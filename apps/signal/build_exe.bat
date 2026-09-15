@echo off
setlocal EnableExtensions
cd /d "%~dp0"
for %%I in ("%~dp0..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"

python -m pip install -r requirements-build.txt
if errorlevel 1 goto :error
python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --name "SGNL_Platform" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%PROJECT_ROOT%" ^
  SIGNAL.py
if errorlevel 1 goto :error

if exist "dist\SGNL_Platform\assets" rmdir /s /q "dist\SGNL_Platform\assets"
xcopy /E /I /Y "%ASSETS%" "dist\SGNL_Platform\assets" >nul
xcopy /E /I /Y "%PROJECT_ROOT%\shared" "dist\SGNL_Platform\shared" >nul

echo.
echo Done: dist\SGNL_Platform\SGNL_Platform.exe
echo Shared images: dist\SGNL_Platform\assets
pause
exit /b 0

:error
echo.
echo Build failed.
pause
exit /b 1
