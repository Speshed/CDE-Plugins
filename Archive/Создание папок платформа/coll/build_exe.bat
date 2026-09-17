@echo off
setlocal EnableExtensions

set "ROOT=%~dp0"
for %%I in ("%~dp0..\..\..") do set "PROJECT_ROOT=%%~fI"
set "ASSETS=%PROJECT_ROOT%\assets"
set "APP_NAME=LarixCommentImporter"

tasklist /FI "IMAGENAME eq %APP_NAME%.exe" /NH | findstr /I /C:"%APP_NAME%.exe" >nul
if not errorlevel 1 (
  echo ERROR: Close %APP_NAME%.exe before rebuilding.
  pause
  exit /b 1
)

if exist "%ROOT%build" rmdir /S /Q "%ROOT%build"
if exist "%ROOT%dist" rmdir /S /Q "%ROOT%dist"
if exist "%ROOT%%APP_NAME%.spec" del /F /Q "%ROOT%%APP_NAME%.spec"

if not exist "%ASSETS%\icon.ico" (
  echo ERROR: Shared assets folder not found: %ASSETS%
  pause
  exit /b 1
)

echo Building new single-file application...
py -m PyInstaller "%ROOT%larix_comment_importer.py" ^
  --onefile ^
  --clean ^
  --noconsole ^
  --noconfirm ^
  --name "%APP_NAME%" ^
  --icon "%ASSETS%\icon.ico"

if errorlevel 1 goto :error
if not exist "%ROOT%dist\%APP_NAME%.exe" goto :error

echo.
echo Build completed:
echo "%ROOT%dist\%APP_NAME%.exe"
echo.
echo UI images are embedded in the importer; no local icon folder is required.
pause
exit /b 0

:error
echo ERROR: Build failed.
pause
exit /b 1
