@echo off
echo Stopping Zawad Proposal Agent...

:: Kill any running pythonw or py process running the agent
taskkill /f /fi "IMAGENAME eq wscript.exe" /fi "WINDOWTITLE eq launcher*" >nul 2>&1
wmic process where "CommandLine like '%%agent_background%%'" delete >nul 2>&1

echo Done. Agent stopped for this session.
echo (It will start again automatically next time you log in.)
echo.
echo To PERMANENTLY remove autostart, run remove_autostart.bat
pause
