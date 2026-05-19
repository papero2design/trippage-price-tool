@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found.
    echo Install Python 3.12 from https://www.python.org/downloads/
    echo Check "Add python.exe to PATH" during installation.
    pause
    exit /b 1
)

if exist "venv\" goto install
python -m venv venv
if %errorlevel% neq 0 (
    echo Failed to create virtual environment.
    pause
    exit /b 1
)

:install
echo Installing packages...
venv\Scripts\python.exe -m pip install --upgrade pip --quiet
venv\Scripts\python.exe -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Package installation failed.
    pause
    exit /b 1
)

echo.
echo Setup complete. Run silhaeng.bat to start.
pause
