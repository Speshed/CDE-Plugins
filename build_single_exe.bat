@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "CDE_BUILD_ALL=1"
set "LOG=%~dp0build\logs\build_single_exe.log"

where py >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=py"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. Install Python and add it to PATH.
        goto :error
    )
    set "PYTHON=python"
)

%PYTHON% -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller is not installed.
    echo Run install_dependencies.bat first.
    goto :error
)

echo ============================================
echo   Larix CDE - build SINGLE portable EXE
echo ============================================
echo.

echo [1/4] Checking source project...
%PYTHON% "%~dp0tools\verify_project.py"
if errorlevel 1 goto :error

echo.
echo [2/4] Packing source modules into one Larix_CDE.exe...
if not exist "%~dp0build\logs" mkdir "%~dp0build\logs"
if not exist "%~dp0release" mkdir "%~dp0release"
if exist "%~dp0build\portable-work" rmdir /s /q "%~dp0build\portable-work"

rem Run PyInstaller through a Python helper. Arguments are passed as a Python
rem list instead of a cmd.exe command line, so launcher.py cannot be lost by
rem batch quoting or line continuation.
%PYTHON% "%~dp0tools\build_single_exe.py" > "%LOG%" 2>&1
if errorlevel 1 (
    echo [ERROR] PyInstaller failed while packing Larix_CDE.exe.
    echo Build log: %LOG%
    echo.
    type "%LOG%"
    goto :error
)

echo.
echo [3/4] Checking portable executable...
%PYTHON% "%~dp0tools\verify_single_exe.py"
if errorlevel 1 goto :error

echo.
echo [4/4] Smoke-testing menu and all plugin modes...
%PYTHON% "%~dp0tools\smoke_portable.py"
if errorlevel 1 goto :error

echo.
echo ============================================
echo   BUILD OK
echo   Send ONLY this file:
echo   release\Larix_CDE.exe
echo ============================================
echo Build log: %LOG%
pause
exit /b 0

:error
echo.
echo [ERROR] Portable build stopped. See the error above.
pause
exit /b 1
