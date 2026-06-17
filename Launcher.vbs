' Launches Create_poster_Pics.ps1 hidden (no PowerShell window).
' Edit ScriptPath if you keep the .ps1 somewhere other than %USERPROFILE%\WildPosting.
Set objShell = CreateObject("WScript.Shell")
UserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
ScriptPath  = UserProfile & "\WildPosting\Create_poster_Pics.ps1"
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ScriptPath & """", 0, False
