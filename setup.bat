@echo off
echo ========================================
echo   Weekly Report Generator - Setup
echo ========================================
echo.

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed.
    echo.
    echo Please install Python 3.7 or higher and try again.
    echo Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/4] Checking Python version...
python --version
echo.

REM Check if virtual environment exists
if exist .venv (
    echo [INFO] Existing virtual environment found.
    set /p "reinstall=Delete and reinstall? (Y/N): "
    if /i "%reinstall%"=="Y" (
        echo [2/4] Removing existing virtual environment...
        rmdir /s /q .venv
        echo [2/4] Creating virtual environment...
        python -m venv .venv
    ) else (
        echo [2/4] Using existing virtual environment...
    )
) else (
    echo [2/4] Creating virtual environment...
    python -m venv .venv
)
echo.

REM Activate virtual environment
echo [3/4] Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)
echo.

REM Upgrade pip
echo [4/4] Installing required packages...
python -m pip install --upgrade pip --quiet
echo   - pip upgraded

REM Install packages from requirements.txt
if exist requirements.txt (
    pip install -r requirements.txt --quiet
    echo   - openpyxl installed
    echo   - pywin32 installed
) else (
    pip install openpyxl pywin32 --quiet
    echo   - openpyxl installed
    echo   - pywin32 installed
)
echo.

echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo To generate reports:
echo.
echo   Manual (interactive):
echo     create.bat
echo.
echo   Automated (scheduled):
echo     create_unattended.bat 20270115
echo.
echo To activate virtual environment:
echo     .venv\Scripts\activate
echo.
pause
