' VBScript
' Purpose: Elevate itself to run commands with administrative privileges
' Usage: Place the commands you want to run as an admin inside the Elevated code block

' Function to check if script is running as admin
Function IsAdmin()
    Dim shell, result
    Set shell = CreateObject("WScript.Shell")
    On Error Resume Next
    shell.RegRead "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" ' try reading a protected key
    If Err.Number = 0 Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
    On Error Goto 0
End Function

' Main execution
If Not IsAdmin() Then
    ' Relaunch the script as admin using ShellExecute
    Dim shell, myScript
    Set shell = CreateObject("Shell.Application")
    myScript = WScript.ScriptFullName
    ' 1 = run as admin, "" = default options
    shell.ShellExecute "wscript.exe", """" & myScript & """", "", "runas", 1
    WScript.Quit
End If

' ===== Elevation successful, place any code below =====
' Example: creating a file in System32 (requires elevated privileges)
Dim fso, filePath, file
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = "C:\Windows\System32\test_admin.txt"

If Not fso.FileExists(filePath) Then
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine("This file was created with admin privileges.")
    file.Close
End If

MsgBox "Script is running with admin privileges. File created successfully."
