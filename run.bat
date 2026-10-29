@echo off
echo Starting Zawad Proposal Agent...

:: Try 'py' launcher first (Windows standard), then 'python'
py --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python not found. Please run install.bat first.
        pause
        exit /b 1
    )
    set PYTHON=python
) else (
    set PYTHON=py
)

:: Auto-install packages if missing
%PYTHON% -c "import watchdog" >nul 2>&1
if errorlevel 1 (
    echo First run: installing required packages...
    %PYTHON% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Installation failed. Please run install.bat manually.
        pause
        exit /b 1
    )
)

cd /d "%~dp0"
%PYTHON% agent.py
