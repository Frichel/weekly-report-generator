@echo off
cd /d "%~dp0"

echo ========================================
echo   주간 업무보고서 생성 (수동 실행)
echo ========================================
echo.

REM 가상환경 확인
if not exist .venv (
    echo [오류] 가상환경이 설치되지 않았습니다.
    echo.
    echo setup.bat를 먼저 실행하여 설치를 완료해주세요.
    echo.
    pause
    exit /b 1
)

REM 가상환경 Python으로 스크립트 실행
.venv\Scripts\python.exe res\weekly-report-generator.py

REM 오류 확인
if errorlevel 1 (
    echo.
    echo [오류] 프로그램 실행 중 오류가 발생했습니다.
    pause
    exit /b 1
)
