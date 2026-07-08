@echo off
echo ========================================================
echo TBRL Smart Noting - Windows Offline Installer
echo ========================================================
echo.

:: 1. Check if python is installed
where python >nvl 2>nul
if %errorlevel% neq 0 (
    python --version >nul 2>nul
    if %errorlevel% neq 0 (
        echo [ERROR] Python is not installed or not in PATH!
        echo Please install Python 3.12 (64-bit) on this PC first.
        echo Make sure to check the box "Add Python to PATH" during installation.
        echo.
        pause
        exit /b
    )
)

:: 2. Create virtual environment
echo Creating virtual environment (venv)...
python -m venv venv
if %errorlevel% neq 0 (
    echo [ERROR] Failed to create virtual environment!
    pause
    exit /b
)

:: 3. Activate venv and install libraries offline
echo.
echo Installing dependencies offline from windows_required_libraries/...
call venv\Scripts\activate.bat

python -m pip install --no-index --find-links=windows_required_libraries -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] Direct installation failed, attempting fallback loop...
    :: Fallback to installing wheels individually in resolved order (dependencies first)
    python -m pip install --no-index --find-links=windows_required_libraries colorama blinker markupsafe typing_extensions lxml click itsdangerous jinja2 werkzeug flask python_docx
)

if %errorlevel% neq 0 (
    echo [ERROR] Offline installation failed!
    pause
    exit /b
)

echo.
echo ========================================================
echo [SUCCESS] Offline environment setup completed perfectly!
echo To start the application, run:
echo   venv\Scripts\python app.py
echo ========================================================
echo.
pause
