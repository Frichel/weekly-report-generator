@echo off
cd /d "%~dp0"

echo ========================================
echo   Weekly Report Generator (Automated)
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

REM Run with date argument or auto-detect
if "%1"=="" (
    echo No date specified - using this Friday
    .venv\Scripts\python.exe res\weekly-report-generator_unattended.py
) else (
    echo Date: %1
    .venv\Scripts\python.exe res\weekly-report-generator_unattended.py %1
)

REM Check for errors
if errorlevel 1 (
    echo.
    echo [ERROR] An error occurred during execution.
    pause
    exit /b 1
)
