@echo off
echo ============================================================
echo  Zawad Proposal Agent - Installing dependencies
echo ============================================================
echo.

:: Try 'py' launcher first (Windows standard), then 'python'
py --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python is not installed or not in PATH.
        echo Please install Python from https://python.org
        echo Make sure to check "Add Python to PATH" during installation.
        pause
        exit /b 1
    )
    set PYTHON=python
) else (
    set PYTHON=py
)

echo Using Python: %PYTHON%
echo Installing required packages...
%PYTHON% -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Some packages failed to install.
    echo Try running as Administrator or check your internet connection.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Installation complete! Run "run.bat" to start the agent.
echo ============================================================
pause
