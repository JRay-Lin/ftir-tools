@echo off
REM Windows build script for FTIR Tools

echo === FTIR Tools Windows Build Script ===
echo.

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.12+ and try again
    pause
    exit /b 1
)

REM Install build dependencies
echo Installing build dependencies...
pip install -r requirements-build.txt
if %errorlevel% neq 0 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

REM Run the build script
echo.
echo Starting build process...
python build_exe.py

echo.
echo Build process completed!
pause