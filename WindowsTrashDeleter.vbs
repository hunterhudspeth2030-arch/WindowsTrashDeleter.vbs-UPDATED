' VBScript
' Purpose: Elevate itself to run commands with administrative privileges
' Usage: Place the commands you want to run as admin inside the Elevated code block

' Function to check if script is running as admin
Function IsAdmin()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    On Error Resume Next
    ' Try reading a protected registry key
    shell.RegRead "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
    If Err.Number = 0 Then
        IsAdmin = True
    Else
        IsAdmin = False
    End If
    On Error GoTo 0
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
' Example: creating a file in a temporary directory (avoiding permission issues)
Dim fso, filePath, file
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a file in a temp directory instead of System32
filePath = "C:\Temp\test_admin.txt"

' Ensure the Temp directory exists
If Not fso.FolderExists("C:\Temp") Then
    fso.CreateFolder("C:\Temp")
End If

If Not fso.FileExists(filePath) Then
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine("This file was created with admin privileges.")
    file.Close
End If

MsgBox "Script is running with admin privileges. File created at: " & filePath
