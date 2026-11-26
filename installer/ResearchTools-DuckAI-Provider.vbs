Option Explicit
Const DUCKAI_API_URL = "http://127.0.0.1:18080/v1/models"

Dim shell, fso, env, rootDir, runtimeDir, duckExe
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set env = shell.Environment("Process")

If WScript.Arguments.Count > 0 Then
    If LCase(WScript.Arguments(0)) = "--self-test" Then WScript.Quit 0
End If

rootDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
runtimeDir = rootDir & "\DuckAI"
duckExe = runtimeDir & "\duck2api-main.exe"

If Not fso.FileExists(duckExe) Then WScript.Quit 0
If ApiReady() Then WScript.Quit 0

If Not ProcessExistsAtPath(duckExe) Then
    env("SERVER_HOST") = "127.0.0.1"
    env("SERVER_PORT") = "18080"
    shell.Run Chr(34) & duckExe & Chr(34), 0, False
    env("SERVER_HOST") = ""
    env("SERVER_PORT") = ""
End If

For attempt = 1 To 30
    WScript.Sleep 2000
    If ApiReady() Then Exit For
Next
WScript.Quit 0

Function ProcessExistsAtPath(executablePath)
    Dim service, processes, process
    ProcessExistsAtPath = False
    On Error Resume Next
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='duck2api-main.exe'")
    If Err.Number = 0 Then
        For Each process In processes
            If LCase(CStr(process.ExecutablePath)) = LCase(CStr(executablePath)) Then
                ProcessExistsAtPath = True
                Exit For
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function ApiReady()
    Dim request, responseBody
    ApiReady = False
    On Error Resume Next
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.SetTimeouts 1000, 1000, 1000, 2500
    request.Open "GET", DUCKAI_API_URL, False
    request.Send
    If Err.Number = 0 Then
        If request.Status = 200 Then
            responseBody = request.ResponseText
            ApiReady = (InStr(1, responseBody, """data""", vbTextCompare) > 0 And _
                InStr(1, responseBody, """id""", vbTextCompare) > 0)
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function
