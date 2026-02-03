@echo off
echo ========================================
echo   주간 업무보고서 자동 생성 도구 설치
echo ========================================
echo.

REM Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo Python 3.7 이상을 설치한 후 다시 실행해주세요.
    echo 다운로드: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [1/4] Python 버전 확인...
python --version
echo.

REM 가상환경 존재 확인
if exist .venv (
    echo [알림] 기존 가상환경이 발견되었습니다.
    set /p "reinstall=기존 가상환경을 삭제하고 재설치하시겠습니까? (Y/N): "
    if /i "%reinstall%"=="Y" (
        echo [2/4] 기존 가상환경 삭제 중...
        rmdir /s /q .venv
        echo [2/4] 가상환경 생성 중...
        python -m venv .venv
    ) else (
        echo [2/4] 기존 가상환경 사용...
    )
) else (
    echo [2/4] 가상환경 생성 중...
    python -m venv .venv
)
echo.

REM 가상환경 활성화
echo [3/4] 가상환경 활성화 중...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [오류] 가상환경 활성화에 실패했습니다.
    pause
    exit /b 1
)
echo.

REM pip 업그레이드
echo [4/4] 필수 패키지 설치 중...
python -m pip install --upgrade pip --quiet
echo   - pip 업그레이드 완료

REM requirements.txt에서 패키지 설치
if exist requirements.txt (
    pip install -r requirements.txt --quiet
    echo   - openpyxl 설치 완료
    echo   - pywin32 설치 완료
) else (
    pip install openpyxl pywin32 --quiet
    echo   - openpyxl 설치 완료
    echo   - pywin32 설치 완료
)
echo.

echo ========================================
echo 설치가 완료되었습니다!
echo ========================================
echo.
echo 다음 명령어로 프로그램을 실행하세요:
echo.
echo   수동 실행 (대화형):
echo     보고서_생성.bat
echo.
echo   자동 실행 (특정 날짜):
echo     보고서_자동생성.bat 20270115
echo.
echo 가상환경을 활성화하려면:
echo     .venv\Scripts\activate
echo.
pause
