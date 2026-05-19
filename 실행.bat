@echo off
chcp 65001 >nul
echo ====================================
echo  트립페이지 가격비교 도구
echo ====================================
echo.

cd /d "%~dp0"

:: 설치 여부 확인
if not exist "venv\Scripts\activate.bat" (
    echo [오류] 설치가 되어 있지 않습니다.
    echo 설치.bat 을 먼저 실행해주세요.
    echo.
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo 잠시 후 브라우저가 자동으로 열립니다...
echo 이 창을 닫으면 프로그램이 종료됩니다.
echo.

streamlit run app.py --server.headless false

pause
