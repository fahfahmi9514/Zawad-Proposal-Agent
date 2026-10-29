@echo off
echo ============================================================
echo  Zawad Proposal Agent — Autostart Setup
echo  Run this ONCE. After this, the agent starts automatically
echo  every time you log into Windows. No windows. No clicking.
echo ============================================================
echo.

set AGENT_DIR=%~dp0
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set DEST=%STARTUP%\ZawadProposalAgent.vbs

:: Copy the silent launcher to Windows Startup folder
copy /y "%AGENT_DIR%launcher.vbs" "%DEST%" >nul

if errorlevel 1 (
    echo ERROR: Could not install to Startup folder.
    echo Path tried: %DEST%
    pause
    exit /b 1
)

echo  Installed to Windows Startup folder.
echo  Location: %DEST%
echo.
echo ============================================================
echo  SUCCESS! Done.
echo.
echo  The agent will now start automatically every time you
echo  log into Windows — silently, no window, no taskbar icon.
echo.
echo  Starting the agent now for this session...
echo ============================================================
echo.

start "" wscript.exe "%AGENT_DIR%launcher.vbs"

echo  Agent is running. Check agent.log to see activity.
echo.
pause
