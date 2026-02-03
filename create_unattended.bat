@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ========================================
echo   주간 업무보고서 생성 (자동 실행)
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

REM 날짜 인수가 있으면 사용, 없으면 자동 (이번 주 금요일)
if "%1"=="" (
    echo 날짜 입력 없음 - 이번 주 금요일로 자동 생성합니다.
    .venv\Scripts\python.exe res\weekly-report-generator_unattended.py
) else (
    echo 날짜: %1
    .venv\Scripts\python.exe res\weekly-report-generator_unattended.py %1
)

REM 오류 확인
if errorlevel 1 (
    echo.
    echo [오류] 프로그램 실행 중 오류가 발생했습니다.
    pause
    exit /b 1
)

echo.
echo 자동 실행 완료!
pause
