' Launches Create_poster_Pics.ps1 hidden (no PowerShell window).
' Edit ScriptPath if you clone the repo somewhere else.
Set objShell = CreateObject("WScript.Shell")
UserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
ScriptPath  = UserProfile & "\OneDrive\Workspace\Projects\WildPosting\Create_poster_Pics.ps1"
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ScriptPath & """", 0, False
