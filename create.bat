@echo off
cd /d "%~dp0"

echo ========================================
echo   Weekly Report Generator (Manual)
echo ========================================
echo.

REM Check virtual environment
if not exist .venv (
    echo [ERROR] Virtual environment not found.
    echo.
    echo Please run setup.bat first.
    echo.
    pause
    exit /b 1
)

REM Run Python script
.venv\Scripts\python.exe res\weekly-report-generator.py

REM Check for errors
if errorlevel 1 (
    echo.
    echo [ERROR] An error occurred during execution.
    pause
    exit /b 1
)
