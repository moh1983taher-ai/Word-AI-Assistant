Option Explicit
' RESEARCHTOOLS_DUCKAI_BOOTSTRAP_V3
Const BOOTSTRAP_SHA256 = "5d92f20d668410d881c3f965e84b625ccc6690e09631c579a372935e8e8aae3a"
Const BOOTSTRAP_RAW_URL = "https://raw.githubusercontent.com/moh1983taher-ai/Word-AI-Assistant/main/runtime/DuckAI-Bootstrap.ps1"
Dim shell,fso,env,rootDir,installDir,launcherPath,bootstrapPath
Dim duckExe,agentExe,bridgeScript,powershellExe,wscriptExe,runtimeReady,lastEnsureAt,ensureCode
Set shell=CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
Set env=shell.Environment("Process")
rootDir=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
installDir=rootDir & "\DuckAI"
launcherPath=rootDir & "\ResearchTools-Launcher.vbs"
bootstrapPath=rootDir & "\DuckAI-Bootstrap.ps1"
duckExe=installDir & "\duck2api-main.exe"
agentExe=installDir & "\ResearchTools-DuckAgent.exe"
bridgeScript=installDir & "\DuckAI-MCP-Bridge.ps1"
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
        If Not runtimeReady Or Not ProcessExists("duck2api-main.exe") Or Not ProcessExists("ResearchTools-DuckAgent.exe") Or Not BridgeListening() Then
            ensureCode=EnsureRuntime(True)
            lastEnsureAt=Now
            runtimeReady=(ensureCode=0)
            If runtimeReady Then StartServices
        End If
        Do While ProcessExists("WINWORD.EXE")
            If Not runtimeReady Or Not ProcessExists("duck2api-main.exe") Or Not ProcessExists("ResearchTools-DuckAgent.exe") Or Not BridgeListening() Then
                If DateDiff("s",lastEnsureAt,Now)>=30 Then
                    ensureCode=EnsureRuntime(False)
                    lastEnsureAt=Now
                    runtimeReady=(ensureCode=0)
                    If runtimeReady Then StartServices
                End If
            End If
            WScript.Sleep 3000
        Loop
        StopBridgeProcess
        StopProcess "ResearchTools-DuckAgent.exe"
        WScript.Sleep 1000
        StopProcess "duck2api-main.exe"
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
    If fso.FileExists(duckExe) Then
        env("SERVER_HOST")="127.0.0.1"
        env("SERVER_PORT")="18080"
        shell.Run Chr(34)&duckExe&Chr(34),0,False
        WScript.Sleep 5000
        env("SERVER_HOST")=""
        env("SERVER_PORT")=""
    End If
    If fso.FileExists(agentExe) Then
        env("SERVER_HOST")=""
        env("SERVER_PORT")=""
        shell.Run Chr(34)&agentExe&Chr(34),0,False
        WScript.Sleep 1500
    End If
    StartBridge
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
        If InStr(commandLine,LCase("DuckAI-MCP-Bridge.ps1"))>0 Then
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
        commandLine=LCase(CStr(process.CommandLine))
        If InStr(commandLine,LCase("DuckAI-MCP-Bridge.ps1"))>0 Then
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
Sub StopProcess(processName)
    On Error Resume Next
    shell.Run "taskkill /F /IM "&processName,0,True
    On Error GoTo 0
End Sub
