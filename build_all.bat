@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ROOT=%~dp0"
set "ASSETS=%ROOT%assets"
set "TMP=%ROOT%_build"
set "PRODUCTS=%TMP%\products"
set "PACKAGE=%ROOT%dist\SOD_Manager"

title Build SOD Manager

echo ============================================
echo   Building SOD Manager package
echo ============================================
echo.

if not exist "%ASSETS%\icon.ico" (
  echo [ERROR] Shared assets folder is missing or incomplete:
  echo   %ASSETS%
  pause
  exit /b 1
)

python -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 goto :error

if exist "%TMP%" rmdir /s /q "%TMP%"
if exist "%ROOT%dist" rmdir /s /q "%ROOT%dist"
mkdir "%PRODUCTS%"
mkdir "%PACKAGE%"

rem ------------------------------------------------------------
rem 1. Larix Platform
rem ------------------------------------------------------------
echo [1/5] Larix Platform...
pushd "%ROOT%apps\larix"
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "Larix_Platform_Plugin" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%ROOT%" ^
  --distpath "%PRODUCTS%\larix" ^
  --workpath "%TMP%\work_larix" ^
  --specpath "%TMP%\spec_larix" ^
  Larix_User_Platform.py
if errorlevel 1 (popd & goto :error)
popd

rem ------------------------------------------------------------
rem 2. VitroCAD
rem ------------------------------------------------------------
echo [2/5] VitroCAD...
pushd "%ROOT%apps\vitrocad"
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "VitroCAD" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%ROOT%" ^
  --distpath "%PRODUCTS%\vitrocad" ^
  --workpath "%TMP%\work_vitrocad" ^
  --specpath "%TMP%\spec_vitrocad" ^
  --hidden-import requests ^
  --hidden-import urllib3 ^
  --hidden-import urllib3.contrib.socks ^
  --hidden-import socks ^
  --hidden-import openpyxl ^
  --hidden-import pandas ^
  VitroCAD.py
if errorlevel 1 (popd & goto :error)
popd

rem ------------------------------------------------------------
rem 3. SIGNAL / SGNL
rem One-folder is intentional because Qt WebEngine is more reliable here.
rem ------------------------------------------------------------
echo [3/5] SIGNAL...
pushd "%ROOT%apps\signal"
python -m PyInstaller --noconfirm --clean --windowed ^
  --name "SGNL_Platform" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%ROOT%" ^
  --distpath "%PRODUCTS%\signal" ^
  --workpath "%TMP%\work_signal" ^
  --specpath "%TMP%\spec_signal" ^
  --collect-all PySide6 ^
  SIGNAL.py
if errorlevel 1 (popd & goto :error)
popd

rem ------------------------------------------------------------
rem 4. Project Point
rem ------------------------------------------------------------
echo [4/5] Project Point...
pushd "%ROOT%apps\projectpoint"
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "ProjectPoint" ^
  --icon "%ASSETS%\icon.ico" ^
  --paths "%ROOT%" ^
  --add-data "projectpoint\resources\builtin_role_permissions.json;projectpoint\resources" ^
  --distpath "%PRODUCTS%\projectpoint" ^
  --workpath "%TMP%\work_projectpoint" ^
  --specpath "%TMP%\spec_projectpoint" ^
  main.py
if errorlevel 1 (popd & goto :error)
popd

rem ------------------------------------------------------------
rem 5. Main launcher
rem All UI images stay external in one shared assets folder.
rem ------------------------------------------------------------
echo [5/5] Main launcher...
pushd "%ROOT%"
python -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "SOD_Manager" ^
  --icon "%ASSETS%\icon.ico" ^
  --distpath "%PRODUCTS%\launcher" ^
  --workpath "%TMP%\work_launcher" ^
  --specpath "%TMP%\spec_launcher" ^
  launcher.py
if errorlevel 1 (popd & goto :error)
popd

rem ------------------------------------------------------------
rem Assemble a single distributable folder.
rem ------------------------------------------------------------
copy /Y "%PRODUCTS%\launcher\SOD_Manager.exe" "%PACKAGE%\SOD_Manager.exe" >nul
xcopy /E /I /Y "%ASSETS%" "%PACKAGE%\assets" >nul
xcopy /E /I /Y "%ROOT%shared" "%PACKAGE%\shared" >nul

mkdir "%PACKAGE%\apps\larix"
copy /Y "%PRODUCTS%\larix\Larix_Platform_Plugin.exe" "%PACKAGE%\apps\larix\Larix_Platform_Plugin.exe" >nul

mkdir "%PACKAGE%\apps\vitrocad"
copy /Y "%PRODUCTS%\vitrocad\VitroCAD.exe" "%PACKAGE%\apps\vitrocad\VitroCAD.exe" >nul

mkdir "%PACKAGE%\apps\signal"
xcopy /E /I /Y "%PRODUCTS%\signal\SGNL_Platform" "%PACKAGE%\apps\signal\SGNL_Platform" >nul

mkdir "%PACKAGE%\apps\projectpoint"
copy /Y "%PRODUCTS%\projectpoint\ProjectPoint.exe" "%PACKAGE%\apps\projectpoint\ProjectPoint.exe" >nul
if exist "%ROOT%apps\projectpoint\connection_profiles.json" copy /Y "%ROOT%apps\projectpoint\connection_profiles.json" "%PACKAGE%\apps\projectpoint\connection_profiles.json" >nul

copy /Y "%ROOT%README.md" "%PACKAGE%\README.md" >nul

echo.
echo ============================================
echo   DONE
echo ============================================
echo Package:
echo   %PACKAGE%
echo.
echo Shared images:
echo   %PACKAGE%\assets
pause
exit /b 0

:error
echo.
echo [ERROR] Build stopped. Check the output above.
pause
exit /b 1
