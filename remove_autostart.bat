@echo off
echo Removing Zawad Proposal Agent from autostart...
set DEST=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\ZawadProposalAgent.vbs
if exist "%DEST%" (
    del /f "%DEST%"
    echo Removed from Startup folder.
) else (
    echo Not found in Startup folder — nothing to remove.
)
wmic process where "CommandLine like '%%agent_background%%'" delete >nul 2>&1
echo Done.
pause
