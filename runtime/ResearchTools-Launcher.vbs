Option Explicit
' RESEARCHTOOLS_DUCKAI_BOOTSTRAP_V1

Const BOOTSTRAP_SHA256 = "09ab9858b3c8cc1739d7b96a713264d2bb1ad87b5e141b9362f745d4f4c291bb"
Const BOOTSTRAP_RAW_URL = "https://raw.githubusercontent.com/moh1983taher-ai/Word-AI-Assistant/main/runtime/DuckAI-Bootstrap.ps1"
Const TASK_NAME = "ResearchTools DuckAI Agent"

Dim shell, fso, env
Dim rootDir, installDir, launcherPath, bootstrapPath
Dim duckExe, agentExe, powershellExe, wscriptExe
Dim runtimeReady, lastEnsureAt, ensureCode

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set env = shell.Environment("Process")

rootDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
installDir = rootDir & "\DuckAI"
launcherPath = rootDir & "\ResearchTools-Launcher.vbs"
bootstrapPath = rootDir & "\DuckAI-Bootstrap.ps1"
duckExe = installDir & "\duck2api-main.exe"
agentExe = installDir & "\ResearchTools-DuckAgent.exe"
powershellExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
wscriptExe = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\wscript.exe"

runtimeReady = False
lastEnsureAt = DateAdd("s", -60, Now)

EnsureFolder rootDir

' Move the durable launcher outside the runtime folder so deleting the EXEs
' or the entire DuckAI subfolder does not remove the recovery entry point.
If LCase(fso.GetAbsolutePathName(WScript.ScriptFullName)) <> LCase(fso.GetAbsolutePathName(launcherPath)) Then
    On Error Resume Next
    fso.CopyFile WScript.ScriptFullName, launcherPath, True
    Err.Clear
    If fso.FileExists(installDir & "\DuckAI-Bootstrap.ps1") Then
        fso.CopyFile installDir & "\DuckAI-Bootstrap.ps1", bootstrapPath, True
    End If
    On Error GoTo 0
    If fso.FileExists(launcherPath) Then
        If EnsureBootstrapScript() Then RunBootstrap True, False
        shell.Run Chr(34) & wscriptExe & Chr(34) & " " & Chr(34) & launcherPath & Chr(34), 0, False
        WScript.Quit
    End If
End If

Do
    If ProcessExists("WINWORD.EXE") Then
        If Not runtimeReady Or Not ProcessExists("duck2api-main.exe") Or Not ProcessExists("ResearchTools-DuckAgent.exe") Then
            ensureCode = EnsureRuntime()
            lastEnsureAt = Now
            runtimeReady = (ensureCode = 0)
        End If

        If runtimeReady Then
            If Not ProcessExists("duck2api-main.exe") Then StartDuck
            If Not ProcessExists("ResearchTools-DuckAgent.exe") Then StartAgent
        End If

        Do While ProcessExists("WINWORD.EXE")
            If Not runtimeReady Or Not ProcessExists("duck2api-main.exe") Or Not ProcessExists("ResearchTools-DuckAgent.exe") Then
                If DateDiff("s", lastEnsureAt, Now) >= 30 Then
                    ensureCode = EnsureRuntime()
                    lastEnsureAt = Now
                    runtimeReady = (ensureCode = 0)
                    If runtimeReady Then
                        If Not ProcessExists("duck2api-main.exe") Then StartDuck
                        If Not ProcessExists("ResearchTools-DuckAgent.exe") Then StartAgent
                    End If
                End If
            End If
            WScript.Sleep 3000
        Loop

        StopProcess "ResearchTools-DuckAgent.exe"
        WScript.Sleep 1000
        StopProcess "duck2api-main.exe"
        runtimeReady = False
    Else
        WScript.Sleep 3000
    End If
Loop

Sub EnsureFolder(path)
    On Error Resume Next
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    On Error GoTo 0
End Sub

Function EnsureBootstrapScript()
    Dim command, result, pathArg, urlArg, hashArg
    EnsureBootstrapScript = False

    If Not fso.FileExists(bootstrapPath) And fso.FileExists(installDir & "\DuckAI-Bootstrap.ps1") Then
        On Error Resume Next
        fso.CopyFile installDir & "\DuckAI-Bootstrap.ps1", bootstrapPath, True
        On Error GoTo 0
    End If

    pathArg = Replace(bootstrapPath, "'", "''")
    urlArg = Replace(BOOTSTRAP_RAW_URL, "'", "''")
    hashArg = BOOTSTRAP_SHA256
    command = "$ErrorActionPreference='Stop';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
        "$p='" & pathArg & "';$u='" & urlArg & "';$h='" & hashArg & "';" & _
        "if(!(Test-Path -LiteralPath $p) -or (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash -ne $h){" & _
        "$t=$p+'.download';Invoke-WebRequest -UseBasicParsing -Uri $u -OutFile $t;" & _
        "if((Get-FileHash -LiteralPath $t -Algorithm SHA256).Hash -ne $h){Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue;exit 51};" & _
        "Move-Item -LiteralPath $t -Destination $p -Force};exit 0"

    result = shell.Run(Chr(34) & powershellExe & Chr(34) & " -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command " & Chr(34) & command & Chr(34), 0, True)
    EnsureBootstrapScript = (result = 0)
End Function

Function RunBootstrap(registerTask, ensureRuntime)
    Dim command, result
    RunBootstrap = 1
    If Not EnsureBootstrapScript() Then Exit Function

    command = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & bootstrapPath & Chr(34)
    If ensureRuntime Then command = command & " -EnsureRuntime"
    If registerTask Then command = command & " -RegisterTask"
    result = shell.Run(Chr(34) & powershellExe & Chr(34) & " " & command, 0, True)
    RunBootstrap = result
End Function

Function EnsureRuntime()
    EnsureRuntime = RunBootstrap(False, True)
End Function

Sub StartDuck()
    If fso.FileExists(duckExe) Then
        env("SERVER_HOST") = "127.0.0.1"
        env("SERVER_PORT") = "18080"
        shell.Run """" & duckExe & """", 0, False
        WScript.Sleep 5000
        env("SERVER_HOST") = ""
        env("SERVER_PORT") = ""
    End If
End Sub

Sub StartAgent()
    If fso.FileExists(agentExe) Then
        env("SERVER_HOST") = ""
        env("SERVER_PORT") = ""
        shell.Run """" & agentExe & """", 0, False
    End If
End Sub

Function ProcessExists(processName)
    Dim service, processes, process
    ProcessExists = False
    On Error Resume Next
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & processName & "'")
    For Each process In processes
        ProcessExists = True
        Exit For
    Next
    On Error GoTo 0
End Function

Sub StopProcess(processName)
    On Error Resume Next
    shell.Run "taskkill /F /IM " & processName, 0, True
    On Error GoTo 0
End Sub