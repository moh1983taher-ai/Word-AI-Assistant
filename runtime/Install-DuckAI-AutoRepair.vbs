Option Explicit

Dim shell, fso, scriptDir, scriptPath, powershellPath, command, exitCode
Dim rootPath, logPath

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

rootPath = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
If Not fso.FolderExists(rootPath) Then fso.CreateFolder rootPath
logPath = fso.BuildPath(rootPath, "DuckAI-AutoRepair-Install.log")

WriteLog "Installer VBS started."

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "Apply-DuckAI-AutoRepair.ps1")
If Not fso.FileExists(scriptPath) Then
    WriteLog "ERROR: Apply-DuckAI-AutoRepair.ps1 is missing beside the VBS file. Extract the complete ZIP before running it."
    WScript.Quit 2
End If

powershellPath = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
If Not fso.FileExists(powershellPath) Then
    WriteLog "ERROR: Windows PowerShell executable was not found."
    WScript.Quit 3
End If

command = Chr(34) & powershellPath & Chr(34) & " -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & scriptPath & Chr(34)
WriteLog "Starting the PowerShell installer silently."
exitCode = shell.Run(command, 0, True)
WriteLog "PowerShell installer finished. ExitCode=" & CStr(exitCode)
WScript.Quit exitCode

Sub WriteLog(message)
    On Error Resume Next
    Dim stream
    Set stream = fso.OpenTextFile(logPath, 8, True)
    stream.WriteLine CStr(Now) & " " & message
    stream.Close
    On Error GoTo 0
End Sub
