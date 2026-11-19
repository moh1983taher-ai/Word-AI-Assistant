# RESEARCHTOOLS_DUCKAI_AUTOREPAIR_INSTALLER_V1
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$root = Join-Path $env:LOCALAPPDATA 'ResearchTools'
$runtime = Join-Path $root 'DuckAI'
$launcher = Join-Path $root 'ResearchTools-Launcher.vbs'
$bootstrap = Join-Path $root 'DuckAI-Bootstrap.ps1'
$log = Join-Path $root 'DuckAI-AutoRepair-Install.log'
$repo = 'moh1983taher-ai/Word-AI-Assistant'

$files = @(
    [pscustomobject]@{
        Name = 'ResearchTools-Launcher.vbs'
        Url = "https://raw.githubusercontent.com/$repo/main/runtime/ResearchTools-Launcher.vbs"
        Path = $launcher
        Sha256 = 'de986495173078e0552731aa5783c323e86c65dd87cebf274609aa7ab9bf061b'
    },
    [pscustomobject]@{
        Name = 'DuckAI-Bootstrap.ps1'
        Url = "https://raw.githubusercontent.com/$repo/main/runtime/DuckAI-Bootstrap.ps1"
        Path = $bootstrap
        Sha256 = '80ff8ca857e04cb23dce0d778fa32139a5c71b6f508f67e7c0a917b9e728d322'
    }
)

function Write-InstallLog {
    param([string]$Message)
    try {
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Add-Content -LiteralPath $log -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -Encoding UTF8
    } catch { }
}

try {
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    Write-InstallLog 'بدأت أداة التهيئة الأولية.'

    if (Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue) {
        throw 'أغلق Microsoft Word قبل تطبيق التحديث الأولي، ثم شغّل الملف مرة أخرى.'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    foreach ($file in $files) {
        $temp = $file.Path + '.download'
        Write-InstallLog ("تنزيل مكوّن التهيئة: " + $file.Name)
        Invoke-WebRequest -UseBasicParsing -Uri $file.Url -OutFile $temp
        $actual = (Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actual -ne $file.Sha256) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
            throw ("فشل التحقق من بصمة الملف: " + $file.Name)
        }
        Move-Item -LiteralPath $temp -Destination $file.Path -Force
    }

    # Stop only the old ResearchTools launcher instances; leave unrelated
    # Windows Script Host processes untouched.
    $oldPaths = @(
        (Join-Path $runtime 'ResearchTools-Launcher.vbs'),
        $launcher
    ) | ForEach-Object { $_.ToLowerInvariant() }

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
    $bootstrapArgs = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $bootstrap + '" -RegisterTask'
    $process = Start-Process -FilePath $powershell -ArgumentList $bootstrapArgs -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw ("تعذر تسجيل مهمة التشغيل التلقائي (ExitCode=" + $process.ExitCode + ').')
    }

    $registeredTask = Get-ScheduledTask -TaskName 'ResearchTools DuckAI Agent' -ErrorAction Stop
    if ($registeredTask.State -eq 'Disabled') {
        throw 'تم إنشاء مهمة التشغيل التلقائي لكنها معطلة.'
    }
    $registeredAction = @($registeredTask.Actions)[0]
    if (-not $registeredAction -or [string]$registeredAction.Execute -notmatch '(?i)wscript\.exe
    Start-Process -FilePath $wscript -ArgumentList ('"' + $launcher + '"') -WindowStyle Hidden
    Write-InstallLog 'اكتمل إعداد المشغّل الدائم. سيجري فحص DuckAI وتنزيله عند أول فتح Word إذا كان مفقوداً أو غير سليم.'
    exit 0
}
catch {
    Write-InstallLog ('فشل التثبيت الأولي: ' + $_.Exception.Message + ' | ' + $_.ScriptStackTrace)
    exit 1
} -or [string]$registeredAction.Arguments -notmatch 'ResearchTools-Launcher\.vbs') {
        throw 'لم يجتز إجراء مهمة التشغيل التلقائي فحص التحقق.'
    }
    Write-InstallLog 'تم التحقق من تسجيل مهمة التشغيل التلقائي وإعداد مشغّلها.'

    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'
    Start-Process -FilePath $wscript -ArgumentList ('"' + $launcher + '"') -WindowStyle Hidden
    Write-InstallLog 'اكتمل إعداد المشغّل الدائم. سيجري فحص DuckAI وتنزيله عند أول فتح Word إذا كان مفقوداً أو غير سليم.'
    exit 0
}
catch {
    Write-InstallLog ('فشل التثبيت الأولي: ' + $_.Exception.Message)
    exit 1
}