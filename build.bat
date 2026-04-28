@echo off
REM Build script for attendance_management.exe
REM Run this script on Windows to produce dist\attendance_management.exe

echo Installing build dependencies...
pip install -r requirements.txt
pip install -r requirements-build.txt

echo.
echo Building executable...
pyinstaller attendance_management.spec

echo.
if exist dist\attendance_management.exe (
    echo Build succeeded: dist\attendance_management.exe
) else (
    echo Build failed. Check the output above for errors.
    exit /b 1
)
