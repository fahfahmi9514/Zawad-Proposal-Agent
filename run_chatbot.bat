@echo off
title Zawad Proposal Chatbot
cd /d "%~dp0"

set PYTHONHOME=
set PYTHONPATH=
set PYTHON=%~dp0python_runtime\python.exe

echo  Stopping old server on port 5050 (if any)...
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":5050 "') do (
    taskkill /F /PID %%a >nul 2>&1
)
ping -n 2 127.0.0.1 >nul

echo.
echo  =============================================
echo   Zawad Proposal Chatbot
echo   http://localhost:5050
echo  =============================================
echo.

"%PYTHON%" chatbot.py

echo.
echo  Server stopped. Press any key to close.
pause
