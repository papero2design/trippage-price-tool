@echo off
chcp 65001 >nul
echo ====================================
echo  트립페이지 가격비교 도구 - 설치
echo ====================================
echo.

cd /d "%~dp0"

:: Python 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo 아래 주소에서 Python 3.12를 설치하세요:
    echo   https://www.python.org/downloads/
    echo.
    echo 설치 시 "Add python.exe to PATH" 체크박스를 반드시 선택하세요.
    echo 설치 완료 후 이 파일을 다시 실행하세요.
    echo.
    pause
    exit /b 1
)

python --version
echo.

:: 가상환경 생성
if exist "venv\" (
    echo 기존 가상환경이 있습니다. 패키지만 업데이트합니다.
) else (
    echo 가상환경 생성 중...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [오류] 가상환경 생성에 실패했습니다.
        pause
        exit /b 1
    )
)

:: 패키지 설치
echo 패키지 설치 중... (처음에는 시간이 걸릴 수 있습니다)
echo.
call venv\Scripts\activate.bat
pip install --upgrade pip --quiet
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo [오류] 패키지 설치에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo ====================================
echo  설치 완료!
echo  이제 실행.bat 을 더블클릭하여 시작하세요.
echo ====================================
echo.
pause
