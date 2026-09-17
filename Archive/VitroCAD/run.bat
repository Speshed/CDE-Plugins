@echo off
cd /d "%~dp0"

where py >nul 2>&1
if not errorlevel 1 (
    py VitroCAD.py
    exit /b %errorlevel%
)

python VitroCAD.py
