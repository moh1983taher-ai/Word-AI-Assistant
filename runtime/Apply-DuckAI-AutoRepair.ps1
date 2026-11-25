# RESEARCHTOOLS_DUCKAI_AUTOREPAIR_INSTALLER_V3
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$root = Join-Path $env:LOCALAPPDATA 'ResearchTools'
$launcher = Join-Path $root 'ResearchTools-Launcher.vbs'
$bootstrap = Join-Path $root 'DuckAI-Bootstrap.ps1'
$log = Join-Path $root 'DuckAI-AutoRepair-Install.log'
$repo = 'moh1983taher-ai/Word-AI-Assistant'

function Write-InstallLog {
    param([string]$Message)
    try {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            New-Item -ItemType Directory -Path $root -Force | Out-Null
        }
        Add-Content -LiteralPath $log -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -Encoding UTF8
    }
    catch { }
}

$files = @(
    [pscustomobject]@{
        Name = 'ResearchTools-Launcher.vbs'
        Url = "https://raw.githubusercontent.com/$repo/main/runtime/ResearchTools-Launcher.vbs"
        Path = $launcher
        Sha256 = 'b29ebc3292202db79e8f170042ceecb2ae800382b823d16fd5cff2565405c424'
    },
    [pscustomobject]@{
        Name = 'DuckAI-Bootstrap.ps1'
        Url = "https://raw.githubusercontent.com/$repo/main/runtime/DuckAI-Bootstrap.ps1"
        Path = $bootstrap
        Sha256 = '3bb7a970b7a6df745322d531b90098229d60f0e92c55e82aa5dc6bc96ab57411'
    }
)

try {
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    Write-InstallLog 'BEGIN: first-time local bootstrap setup.'

    if (Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue) {
        throw 'أغلق Microsoft Word بالكامل قبل تطبيق التهيئة الأولية.'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    foreach ($file in $files) {
        $temp = $file.Path + '.download'
        Write-InstallLog ("Downloading and verifying: " + $file.Name)
        Invoke-WebRequest -UseBasicParsing -TimeoutSec 60 -Uri $file.Url -OutFile $temp
        $actual = (Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash.ToLowerInvariant()

        if ($actual -ne $file.Sha256) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
            throw ("SHA-256 verification failed for " + $file.Name + ". Expected " + $file.Sha256 + "; got " + $actual)
        }

        Move-Item -LiteralPath $temp -Destination $file.Path -Force
        Write-InstallLog ("Verified and installed: " + $file.Name)
    }

    $oldPaths = @(
        (Join-Path $root 'DuckAI-old-backup\ResearchTools-Launcher.vbs'),
        $launcher
    ) | ForEach-Object { ([string]$_).ToLowerInvariant() }

    $oldHosts = @(Get-CimInstance Win32_Process -Filter "Name='wscript.exe'" -ErrorAction SilentlyContinue)
    foreach ($hostProcess in $oldHosts) {
        $commandLine = [string]$hostProcess.CommandLine
        if (-not $commandLine) { continue }
        $normalized = $commandLine.ToLowerInvariant()
        if ($oldPaths | Where-Object { $normalized.Contains($_) }) {
            Stop-Process -Id $hostProcess.ProcessId -Force -ErrorAction SilentlyContinue
        }
    }

    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $powershell -PathType Leaf)) {
        throw 'تعذر العثور على Windows PowerShell.'
    }

    $bootstrapArgs = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $bootstrap + '" -RegisterTask -EnsureRuntime -CheckForUpdates'
    $process = Start-Process -FilePath $powershell -ArgumentList $bootstrapArgs -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw ("Bootstrap failed while registering startup task. ExitCode=" + $process.ExitCode + ". See DuckAI-Bootstrap.log.")
    }

    $registeredTask = Get-ScheduledTask -TaskName 'ResearchTools DuckAI Agent' -ErrorAction Stop
    if ($registeredTask.State -eq 'Disabled') {
        throw 'تم تسجيل مهمة التشغيل التلقائي لكنها معطلة.'
    }

    $actions = @($registeredTask.Actions)
    if ($actions.Count -eq 0) {
        throw 'مهمة التشغيل التلقائي لا تحتوي على إجراء.'
    }

    $registeredAction = $actions[0]
    if ([string]$registeredAction.Execute -notmatch '(?i)wscript\.exe$' -or [string]$registeredAction.Arguments -notmatch 'ResearchTools-Launcher\.vbs') {
        throw 'فشل التحقق من مسار المشغّل المسجل في مهمة ويندوز.'
    }

    Write-InstallLog 'SUCCESS: startup task exists and points to the durable launcher.'
    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'
    Start-Process -FilePath $wscript -ArgumentList ('"' + $launcher + '"') -WindowStyle Hidden
    Write-InstallLog 'SUCCESS: durable launcher started. DuckAI will be checked and restored when Word opens.'
    exit 0
}
catch {
    Write-InstallLog ('ERROR: ' + $_.Exception.Message)
    if ($_.ScriptStackTrace) {
        Write-InstallLog ('STACK: ' + $_.ScriptStackTrace)
    }
    exit 1
}
