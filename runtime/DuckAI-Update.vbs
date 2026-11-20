Option Explicit

Dim shell, fso, rootDir, tempDir, psPath, resultPath, logPath, powershellPath
Dim scriptText, stream, command, exitCode, resultText

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count > 0 Then
    If LCase(WScript.Arguments(0)) = "--self-test" Then WScript.Quit 0
End If

rootDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ResearchTools"
tempDir = shell.ExpandEnvironmentStrings("%TEMP%")
If Not fso.FolderExists(rootDir) Then fso.CreateFolder rootDir
psPath = fso.BuildPath(tempDir, "DuckAI-Update-" & fso.GetTempName & ".ps1")
resultPath = fso.BuildPath(rootDir, "DuckAI-Update-Result.txt")
logPath = fso.BuildPath(rootDir, "DuckAI-Update.log")
powershellPath = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"

scriptText = Join(Array( _
"$ErrorActionPreference = 'Stop'", _
"$ProgressPreference = 'SilentlyContinue'", _
"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12", _
"$repo = 'moh1983taher-ai/Word-AI-Assistant'", _
"$root = Join-Path $env:LOCALAPPDATA 'ResearchTools'", _
"$resultPath = Join-Path $root 'DuckAI-Update-Result.txt'", _
"$logPath = Join-Path $root 'DuckAI-Update.log'", _
"$temp = Join-Path $env:TEMP ('DuckAI-Update-' + [Guid]::NewGuid().ToString('N'))", _
"function Write-UpdateLog([string]$message) { Add-Content -LiteralPath $logPath -Value ((Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' ' + $message) -Encoding UTF8 }", _
"New-Item -ItemType Directory -Path $root -Force | Out-Null", _
"Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue", _
"try {", _
"    if (Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue) { throw 'Close all Microsoft Word windows before updating.' }", _
"    $headers = @{ 'User-Agent' = 'ResearchTools-DuckAI-OneClickUpdater/1.0'; 'Accept' = 'application/vnd.github+json' }", _
"    $release = Invoke-RestMethod -TimeoutSec 60 -Uri ('https://api.github.com/repos/' + $repo + '/releases/latest') -Headers $headers -Method Get", _
"    if ([string]$release.tag_name -notmatch '^duckai-runtime-v\d+\.\d+\.\d+$' -or $release.draft -or $release.prerelease) { throw 'No valid stable DuckAI release was found.' }", _
"    $asset = @($release.assets | Where-Object { $_.name -eq 'DuckAI-AutoRepair-Installer.zip' } | Select-Object -First 1)", _
"    if ($asset.Count -eq 0) { throw 'The updater package is missing from the latest release.' }", _
"    $asset = $asset[0]", _
"    if ([string]$asset.browser_download_url -notmatch '^https://github\.com/moh1983taher-ai/Word-AI-Assistant/releases/download/' -or [int64]$asset.size -lt 1000 -or [int64]$asset.size -gt 20MB) { throw 'Published updater package metadata is invalid.' }", _
"    if ([string]$asset.digest -notmatch '^sha256:([A-Fa-f0-9]{64})$') { throw 'GitHub did not provide a SHA-256 digest for the updater package.' }", _
"    $expected = $matches[1].ToUpperInvariant()", _
"    New-Item -ItemType Directory -Path $temp -Force | Out-Null", _
"    $zip = Join-Path $temp 'DuckAI-AutoRepair-Installer.zip'", _
"    Write-UpdateLog ('Downloading updater package for ' + $release.tag_name)", _
"    Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 -Uri $asset.browser_download_url -Headers $headers -OutFile $zip", _
"    if ((Get-Item -LiteralPath $zip).Length -ne [int64]$asset.size) { throw 'Updater package size does not match GitHub metadata.' }", _
"    if ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToUpperInvariant() -ne $expected) { throw 'Updater package SHA-256 verification failed. Installation was not started.' }", _
"    $stage = Join-Path $temp 'installer'", _
"    Expand-Archive -LiteralPath $zip -DestinationPath $stage -Force", _
"    $installer = Join-Path $stage 'Install-DuckAI-AutoRepair.vbs'", _
"    $apply = Join-Path $stage 'Apply-DuckAI-AutoRepair.ps1'", _
"    if (-not (Test-Path -LiteralPath $installer -PathType Leaf) -or -not (Test-Path -LiteralPath $apply -PathType Leaf)) { throw 'The updater package does not contain the required installer files.' }", _
"    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'", _
"    $process = Start-Process -FilePath $wscript -ArgumentList ('"' + $installer + '"') -WindowStyle Hidden -Wait -PassThru", _
"    if ($process.ExitCode -ne 0) { throw ('Local installation failed with exit code ' + $process.ExitCode + '. See DuckAI-AutoRepair-Install.log.') }", _
"    [System.IO.File]::WriteAllText($resultPath, ('SUCCESS|' + $release.tag_name + '|' + 'DuckAI runtime, launcher and MCP bridge were checked and updated.'), [System.Text.UTF8Encoding]::new($true))", _
"    Write-UpdateLog ('Successfully installed ' + $release.tag_name)", _
"} catch {", _
"    $message = [string]$_.Exception.Message", _
"    [System.IO.File]::WriteAllText($resultPath, ('ERROR|' + $message), [System.Text.UTF8Encoding]::new($true))", _
"    Write-UpdateLog ('ERROR: ' + $message)", _
"    exit 1", _
"} finally {", _
"    if (Test-Path -LiteralPath $temp -PathType Container) { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }", _
"}" _
), [Environment]::NewLine)

Set stream = fso.CreateTextFile(psPath, True, True)
stream.Write scriptText
stream.Close

If Not fso.FileExists(powershellPath) Then
    MsgBox "Windows PowerShell was not found.", vbCritical, "DuckAI Update"
    WScript.Quit 3
End If

command = Chr(34) & powershellPath & Chr(34) & " -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & psPath & Chr(34)
exitCode = shell.Run(command, 0, True)

resultText = "Could not determine the update result. Exit code: " & CStr(exitCode)
If fso.FileExists(resultPath) Then resultText = ReadUtf8(resultPath)

On Error Resume Next
If fso.FileExists(psPath) Then fso.DeleteFile psPath, True
On Error GoTo 0

If exitCode = 0 And Left(resultText, 8) = "SUCCESS|" Then
    MsgBox Mid(resultText, 9) & vbCrLf & vbCrLf & "Reopen Microsoft Word. Log: " & logPath, vbInformation, "DuckAI Update"
Else
    MsgBox "The update did not complete." & vbCrLf & resultText & vbCrLf & vbCrLf & "Log: " & logPath, vbExclamation, "DuckAI Update"
End If
WScript.Quit exitCode

Function ReadUtf8(path)
    Dim fileStream
    ReadUtf8 = ""
    On Error Resume Next
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2
    fileStream.Charset = "utf-8"
    fileStream.Open
    fileStream.LoadFromFile path
    ReadUtf8 = fileStream.ReadText
    fileStream.Close
    Set fileStream = Nothing
    On Error GoTo 0
End Function
