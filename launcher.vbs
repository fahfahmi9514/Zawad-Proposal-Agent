' launcher.vbs — runs agent_background.py silently (no window, no taskbar icon)
' This file is called by Task Scheduler at every Windows login.
Dim agentDir
agentDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
CreateObject("WScript.Shell").Run "py """ & agentDir & "agent_background.py""", 0, False
