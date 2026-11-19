Option Explicit
Dim shell, fso, scriptDir, scriptPath, powershellPath, command, exitCode
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "Apply-DuckAI-AutoRepair.ps1")
powershellPath = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
If Not fso.FileExists(scriptPath) Then WScript.Quit 2
command = Chr(34) & powershellPath & Chr(34) & " -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & scriptPath & Chr(34)
exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode