Option Explicit
' RESEARCHTOOLS_DUCKAI_BOOTSTRAP_V3
' RESEARCHTOOLS_MCPHOST_V1
Const BOOTSTRAP_SHA256 = "59f0a32ad419d60b0d95ab6d591ac971e7b5a7e07ab19937106ebe09a500281f"
Const BOOTSTRAP_RAW_URL = "https://raw.githubusercontent.com/moh1983taher-ai/Word-AI-Assistant/main/runtime/DuckAI-Bootstrap.ps1"
Dim shell,fso,env,rootDir,installDir,launcherPath,bootstrapPath
Dim duckExe,bridgeScript,mcpHostExe,powershellExe,wscriptExe,runtimeReady,lastEnsureAt,ensureCode
Set shell=CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
Set env=shell.Environment("Process")

If WScript.Arguments.Count > 0 Then
    If LCase(WScript.Arguments(0)) = "--self-test" Then WScript.Quit 0
End If
rootDir=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
installDir=rootDir & "\DuckAI"
launcherPath=rootDir & "\ResearchTools-Launcher.vbs"
bootstrapPath=rootDir & "\DuckAI-Bootstrap.ps1"
duckExe=installDir & "\duck2api-main.exe"
bridgeScript=installDir & "\DuckAI-MCP-Bridge.ps1"
mcpHostExe=rootDir & "\McpHost\ResearchTools.McpHost.exe"
powershellExe=shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
wscriptExe=shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\wscript.exe"
runtimeReady=False
lastEnsureAt=DateAdd("s",-60,Now)
EnsureFolder rootDir

If LCase(fso.GetAbsolutePathName(WScript.ScriptFullName)) <> LCase(fso.GetAbsolutePathName(launcherPath)) Then
    On Error Resume Next
    fso.CopyFile WScript.ScriptFullName,launcherPath,True
    Err.Clear
    On Error GoTo 0
    If Not fso.FileExists(launcherPath) Then
        WriteLauncherLog "ERROR: could not copy durable launcher."
        WScript.Quit 10
    End If
    If Not EnsureBootstrapScript() Then
        WriteLauncherLog "ERROR: bootstrap verification failed."
        WScript.Quit 11
    End If
    ensureCode=RunBootstrap(True,False,False)
    If ensureCode <> 0 Then
        WriteLauncherLog "ERROR: startup task registration failed. ExitCode=" & CStr(ensureCode)
        WScript.Quit 12
    End If
    shell.Run Chr(34)&wscriptExe&Chr(34)&" "&Chr(34)&launcherPath&Chr(34),0,False
    WScript.Quit
End If

Do
    If ProcessExists("WINWORD.EXE") Then
        If Not runtimeReady Or Not ProcessExistsAtPath(duckExe) Or Not ResearchMcpRunning() Then
            ensureCode=EnsureRuntime(True)
            lastEnsureAt=Now
            runtimeReady=(ensureCode=0)
            If runtimeReady Then
                StartServices
                runtimeReady=(DuckApiReady() And ResearchMcpRunning())
            End If
        End If
        Do While ProcessExists("WINWORD.EXE")
            If Not runtimeReady Or Not ProcessExistsAtPath(duckExe) Or Not ResearchMcpRunning() Then
                If DateDiff("s",lastEnsureAt,Now)>=30 Then
                    ensureCode=EnsureRuntime(False)
                    lastEnsureAt=Now
                    runtimeReady=(ensureCode=0)
                    If runtimeReady Then
                StartServices
                runtimeReady=(DuckApiReady() And ResearchMcpRunning())
            End If
                End If
            End If
            WScript.Sleep 3000
        Loop
        StopResearchMcp
        WScript.Sleep 1000
        StopProcessAtPath duckExe
        runtimeReady=False
    Else
        WScript.Sleep 3000
    End If
Loop

Sub EnsureFolder(path)
    On Error Resume Next
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    On Error GoTo 0
End Sub
Sub WriteLauncherLog(message)
    On Error Resume Next
    Dim stream,logPath
    logPath=fso.BuildPath(rootDir,"DuckAI-Bootstrap.log")
    Set stream=fso.OpenTextFile(logPath,8,True)
    stream.WriteLine CStr(Now)&" "&message
    stream.Close
    On Error GoTo 0
End Sub
Function EnsureBootstrapScript()
    Dim command,result,pathArg,urlArg
    EnsureBootstrapScript=False
    pathArg=Replace(bootstrapPath,"'","''")
    urlArg=Replace(BOOTSTRAP_RAW_URL,"'","''")
    command="$ErrorActionPreference='Stop';[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" & _
        "$p='"&pathArg&"';$u='"&urlArg&"';$h='"&BOOTSTRAP_SHA256&"';" & _
        "if(!(Test-Path -LiteralPath $p) -or (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLowerInvariant() -ne $h){" & _
        "$t=$p+'.download';Invoke-WebRequest -UseBasicParsing -Uri $u -OutFile $t;" & _
        "if((Get-FileHash -LiteralPath $t -Algorithm SHA256).Hash.ToLowerInvariant() -ne $h){Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue;exit 51};" & _
        "Move-Item -LiteralPath $t -Destination $p -Force};exit 0"
    result=shell.Run(Chr(34)&powershellExe&Chr(34)&" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "&Chr(34)&command&Chr(34),0,True)
    EnsureBootstrapScript=(result=0)
End Function
Function RunBootstrap(registerTask,ensureRuntime,checkUpdates)
    Dim command,result
    RunBootstrap=1
    If Not EnsureBootstrapScript() Then Exit Function
    command="-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "&Chr(34)&bootstrapPath&Chr(34)
    If ensureRuntime Then command=command&" -EnsureRuntime"
    If registerTask Then command=command&" -RegisterTask"
    If checkUpdates Then command=command&" -CheckForUpdates"
    result=shell.Run(Chr(34)&powershellExe&Chr(34)&" "&command,0,True)
    RunBootstrap=result
End Function
Function EnsureRuntime(checkUpdates)
    EnsureRuntime=RunBootstrap(False,True,checkUpdates)
End Function
Sub StartServices()
    If Not fso.FileExists(duckExe) Then
        WriteLauncherLog "ERROR: Duck2API executable is missing; MCP host startup was skipped."
        Exit Sub
    End If

    If Not ProcessExistsAtPath(duckExe) Then
        If ProcessExists("duck2api-main.exe") Then
            If Not DuckApiReady() Then
                WriteLauncherLog "ERROR: another duck2api-main.exe process is present, but its installed path cannot be verified and 18080 is not ready; refusing to start a duplicate."
                Exit Sub
            End If
            WriteLauncherLog "INFO: an already-running Duck2API endpoint is responding; reusing 127.0.0.1:18080."
        Else
            env("SERVER_HOST")="127.0.0.1"
            env("SERVER_PORT")="18080"
            shell.Run Chr(34)&duckExe&Chr(34),0,False
            env("SERVER_HOST")=""
            env("SERVER_PORT")=""
        End If
    End If

    If Not WaitForDuckApi(120) Then
        WriteLauncherLog "ERROR: Duck2API did not return a valid non-empty /v1/models response on 127.0.0.1:18080 within 120 seconds; MCP host startup was skipped."
        Exit Sub
    End If

    WriteLauncherLog "INFO: using Duck2API directly on 127.0.0.1:18080; the legacy DuckAgent proxy on 8080 is not started."
    StartMcpHost
End Sub
Function WaitForDuckApi(timeoutSeconds)
    Dim startedAt
    WaitForDuckApi=False
    startedAt=Now
    WriteLauncherLog "INFO: waiting for Duck2API /v1/models before starting the MCP host."
    Do While DateDiff("s",startedAt,Now)<timeoutSeconds
        If DuckApiReady() Then
            WriteLauncherLog "SUCCESS: Duck2API /v1/models returned a non-empty model list after " & CStr(DateDiff("s",startedAt,Now)) & " seconds."
            WaitForDuckApi=True
            Exit Function
        End If
        WScript.Sleep 2000
    Loop
End Function

Function DuckApiReady()
    Dim request,status,responseBody
    DuckApiReady=False
    On Error Resume Next
    Set request=CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number<>0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    request.SetTimeouts 2000,2000,2000,12000
    request.Open "GET","http://127.0.0.1:18080/v1/models",False
    request.Send
    status=0
    responseBody=""
    If Err.Number=0 Then
        status=request.Status
        If status=200 Then responseBody=request.ResponseText
    End If
    Err.Clear
    If status=200 Then
        DuckApiReady=(InStr(1,responseBody,"""data""",vbTextCompare)>0 And InStr(1,responseBody,"""id""",vbTextCompare)>0)
    End If
    Set request=Nothing
    On Error GoTo 0
End Function

Sub StartMcpHost()
    Dim attempt
    If Not fso.FileExists(mcpHostExe) Then
        WriteLauncherLog "WARNING: .NET MCP host is missing; using the legacy bridge for compatibility."
        StartBridge
        Exit Sub
    End If

    If McpHostListening() Then
        If BridgeProcessExists() Then StopBridgeProcess
        WriteLauncherLog "SUCCESS: .NET MCP host is already listening on 127.0.0.1:18081."
        Exit Sub
    End If

    If BridgeProcessExists() Then
        WriteLauncherLog "INFO: stopping the legacy MCP bridge before starting the .NET host."
        StopBridgeProcess
        WScript.Sleep 700
    End If
    If ProcessExistsAtPath(mcpHostExe) Then
        WriteLauncherLog "WARNING: the installed .NET MCP host process is present but /ping is not ready; restarting only that verified installation."
        StopMcpHostProcess
        WScript.Sleep 500
        If ProcessExistsAtPath(mcpHostExe) Then
            WriteLauncherLog "ERROR: the installed .NET MCP host process could not be stopped safely; no duplicate host was started."
            Exit Sub
        End If
    ElseIf ProcessExists("ResearchTools.McpHost.exe") Then
        WriteLauncherLog "ERROR: another MCP host process is present but its executable path cannot be verified; refusing to terminate it or start a duplicate."
        Exit Sub
    End If

    shell.Run Chr(34)&mcpHostExe&Chr(34),0,False
    For attempt=1 To 25
        WScript.Sleep 1000
        If McpHostListening() Then
            WriteLauncherLog "SUCCESS: .NET MCP host is listening on 127.0.0.1:18081."
            Exit Sub
        End If
    Next
    WriteLauncherLog "ERROR: .NET MCP host did not become ready on 127.0.0.1:18081. Inspect MCP host log."
End Sub
Function ResearchMcpRunning()
    If fso.FileExists(mcpHostExe) Then
        ResearchMcpRunning=McpHostListening()
    Else
        ResearchMcpRunning=BridgeListening()
    End If
End Function
Function McpHostListening()
    Dim request,status,responseBody
    McpHostListening=False
    On Error Resume Next
    Set request=CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number<>0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    request.SetTimeouts 1000,1000,1000,1000
    request.Open "GET","http://127.0.0.1:18081/ping",False
    request.Send
    status=0
    responseBody=""
    If Err.Number=0 Then
        status=request.Status
        responseBody=CStr(request.ResponseText)
    End If
    Err.Clear
    McpHostListening=(status=200 And InStr(1,responseBody,"ResearchTools.McpHost",vbTextCompare)>0)
    Set request=Nothing
    On Error GoTo 0
End Function
Sub StopResearchMcp()
    StopBridgeProcess
    StopMcpHostProcess
End Sub
Sub StopMcpHostProcess()
    StopProcessAtPath mcpHostExe
End Sub
Sub StartBridge()
    Dim attempt
    If Not fso.FileExists(bridgeScript) Then
        WriteLauncherLog "ERROR: MCP bridge script is missing."
        Exit Sub
    End If
    If BridgeListening() Then Exit Sub
    If BridgeProcessExists() Then
        WriteLauncherLog "WARNING: MCP bridge process exists but port 18081 is not responding; restarting it."
        StopBridgeProcess
        WScript.Sleep 500
    End If
    shell.Run Chr(34)&powershellExe&Chr(34)&" -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "&Chr(34)&bridgeScript&Chr(34),0,False
    For attempt=1 To 10
        WScript.Sleep 1000
        If BridgeListening() Then
            WriteLauncherLog "SUCCESS: MCP bridge is listening on 127.0.0.1:18081."
            Exit Sub
        End If
    Next
    WriteLauncherLog "ERROR: MCP bridge did not start listening on 127.0.0.1:18081. Inspect DuckAI-MCP-Bridge.log."
End Sub
Function BridgeListening()
    Dim request,status
    BridgeListening=False
    On Error Resume Next
    Set request=CreateObject("WinHttp.WinHttpRequest.5.1")
    If Err.Number<>0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    request.SetTimeouts 1000,1000,1000,1000
    request.Open "GET","http://127.0.0.1:18081/ping",False
    request.Send
    status=0
    If Err.Number=0 Then status=request.Status
    Err.Clear
    BridgeListening=(status=200 Or status=503)
    Set request=Nothing
    On Error GoTo 0
End Function
Function BridgeProcessExists()
    Dim service,processes,process,commandLine
    BridgeProcessExists=False
    On Error Resume Next
    Set service=GetObject("winmgmts:\\.\root\cimv2")
    Set processes=service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    For Each process In processes
        commandLine=LCase(CStr(process.CommandLine))
        If InStr(commandLine,LCase(bridgeScript))>0 Then
            BridgeProcessExists=True
            Exit For
        End If
    Next
    On Error GoTo 0
End Function
Sub StopBridgeProcess()
    Dim service,processes,process,commandLine
    On Error Resume Next
    Set service=GetObject("winmgmts:\\.\root\cimv2")
    Set processes=service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='powershell.exe'")
    For Each process In processes
        commandLine=""
        If Not IsNull(process.CommandLine) Then commandLine=LCase(CStr(process.CommandLine))
        If Len(commandLine)>0 And InStr(commandLine,LCase(fso.GetAbsolutePathName(bridgeScript)))>0 Then
            shell.Run "taskkill /F /PID "&CStr(process.ProcessId),0,True
        End If
    Next
    On Error GoTo 0
End Sub
Function ProcessExists(processName)
    Dim service,processes,process
    ProcessExists=False
    On Error Resume Next
    Set service=GetObject("winmgmts:\\.\root\cimv2")
    Set processes=service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='"&processName&"'")
    For Each process In processes
        ProcessExists=True
        Exit For
    Next
    On Error GoTo 0
End Function

Function ProcessExistsAtPath(executablePath)
    Dim service,processes,process
    ProcessExistsAtPath=False
    On Error Resume Next
    Set service=GetObject("winmgmts:\\.\root\cimv2")
    Set processes=service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='"&fso.GetFileName(executablePath)&"'")
    For Each process In processes
        If ProcessPathMatches(process,executablePath) Then
            ProcessExistsAtPath=True
            Exit For
        End If
    Next
    On Error GoTo 0
End Function

Function ProcessPathMatches(process,executablePath)
    Dim currentPath,commandLine,expectedPath
    ProcessPathMatches=False
    expectedPath=LCase(fso.GetAbsolutePathName(executablePath))
    currentPath=""
    commandLine=""
    On Error Resume Next
    If Not IsNull(process.ExecutablePath) Then currentPath=CStr(process.ExecutablePath)
    If Not IsNull(process.CommandLine) Then commandLine=LCase(CStr(process.CommandLine))
    On Error GoTo 0
    If Len(currentPath)>0 Then
        ProcessPathMatches=(LCase(fso.GetAbsolutePathName(currentPath))=expectedPath)
    ElseIf Len(commandLine)>0 Then
        ProcessPathMatches=(InStr(commandLine,expectedPath)>0)
    End If
End Function

Sub StopProcessAtPath(executablePath)
    Dim service,processes,process
    On Error Resume Next
    Set service=GetObject("winmgmts:\\.\root\cimv2")
    Set processes=service.ExecQuery("SELECT * FROM Win32_Process WHERE Name='"&fso.GetFileName(executablePath)&"'")
    For Each process In processes
        If ProcessPathMatches(process,executablePath) Then
            shell.Run "taskkill /F /PID "&CStr(process.ProcessId),0,True
        End If
    Next
    On Error GoTo 0
End Sub
