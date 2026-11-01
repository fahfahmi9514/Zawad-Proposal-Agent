@echo off
title Zawad - Update & Restart
cd /d "%~dp0"

echo.
echo  =============================================
echo   Zawad Proposal Agent - Update
echo  =============================================
echo.
echo  Checking for updates from GitHub...
echo.

git pull origin main

if errorlevel 1 (
    echo.
    echo  ERROR: Could not connect to GitHub.
    echo  Make sure you are connected to the internet.
    echo.
    pause
    exit /b 1
)

echo.
echo  =============================================
echo   Update complete! Starting the agent...
echo  =============================================
echo.
timeout /t 2 /nobreak >nul

call "%~dp0run_chatbot.bat"
