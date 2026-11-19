# RESEARCHTOOLS_DUCKAI_BOOTSTRAP_V3
[CmdletBinding()]
param(
    [switch]$EnsureRuntime,
    [switch]$RegisterTask
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$Repository = 'moh1983taher-ai/Word-AI-Assistant'
$AssetName = 'DuckAI-Runtime-release.zip'
$TaskName = 'ResearchTools DuckAI Agent'
$RootDir = Join-Path $env:LOCALAPPDATA 'ResearchTools'
$RuntimeDir = Join-Path $RootDir 'DuckAI'
$LauncherPath = Join-Path $RootDir 'ResearchTools-Launcher.vbs'
$LogPath = Join-Path $RootDir 'DuckAI-Bootstrap.log'
$ReleaseApi = "https://api.github.com/repos/$Repository/releases?per_page=100"

function Write-Log {
    param([string]$Message)
    try {
        if (-not (Test-Path -LiteralPath $RootDir -PathType Container)) {
            New-Item -ItemType Directory -Path $RootDir -Force | Out-Null
        }
        if ((Test-Path -LiteralPath $LogPath -PathType Leaf) -and (Get-Item -LiteralPath $LogPath).Length -gt 1MB) {
            Move-Item -LiteralPath $LogPath -Destination ($LogPath + '.old') -Force
        }
        Add-Content -LiteralPath $LogPath -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -Encoding UTF8
    }
    catch {
        # Logging must never interrupt a recovery attempt.
    }
}

function Test-RuntimeIntegrity {
    param([Parameter(Mandatory = $true)][string]$Directory)

    $sumFile = Join-Path $Directory 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $sumFile -PathType Leaf)) {
        return $false
    }

    $required = @(
        'BUILD-INFO.txt',
        'ResearchTools-DuckAgent.exe',
        'ResearchTools-Launcher.vbs',
        'DuckAI-Bootstrap.ps1',
        'duck2api-main.exe',
        'manifest.xml',
        'uninstall.ps1'
    )

    foreach ($name in $required) {
        $candidate = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $false
        }
    }

    $entries = @{}
    foreach ($line in (Get-Content -LiteralPath $sumFile -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        if ($line -notmatch '^\s*([A-Fa-f0-9]{64})\s+(.+?)\s*$') {
            return $false
        }

        $expected = $matches[1].ToUpperInvariant()
        $name = $matches[2].Trim()

        if ([System.IO.Path]::GetFileName($name) -ne $name -or $name -match '[\\/:]' -or $name -eq '.' -or $name -eq '..') {
            return $false
        }
        if ($entries.ContainsKey($name)) {
            return $false
        }
        $entries[$name] = $expected
    }

    foreach ($name in $required) {
        if (-not $entries.ContainsKey($name)) {
            return $false
        }
    }

    foreach ($name in $entries.Keys) {
        $candidate = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $false
        }
        try {
            $actual = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToUpperInvariant()
        }
        catch {
            return $false
        }
        if ($actual -ne $entries[$name]) {
            return $false
        }
    }

    return $true
}

function Test-ZipPaths {
    param([Parameter(Mandatory = $true)][string]$ArchivePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
    try {
        $seen = @{}
        if ($archive.Entries.Count -gt 32) {
            return $false
        }

        [int64]$uncompressedBytes = 0
        foreach ($entry in $archive.Entries) {
            $name = $entry.FullName
            if ([string]::IsNullOrWhiteSpace($name) -or $name.Contains('/') -or $name.Contains('\') -or $name.Contains(':')) {
                return $false
            }
            if ($entry.Name -ne $name -or $name -eq '.' -or $name -eq '..') {
                return $false
            }
            if ($seen.ContainsKey($name)) {
                return $false
            }

            $seen[$name] = $true
            $uncompressedBytes += [int64]$entry.Length
            if ($uncompressedBytes -gt 250MB) {
                return $false
            }
        }

        return $true
    }
    finally {
        $archive.Dispose()
    }
}

function Register-StartupTask {
    if (-not (Test-Path -LiteralPath $LauncherPath -PathType Leaf)) {
        $runtimeLauncher = Join-Path $RuntimeDir 'ResearchTools-Launcher.vbs'
        if (Test-Path -LiteralPath $runtimeLauncher -PathType Leaf) {
            Copy-Item -LiteralPath $runtimeLauncher -Destination $LauncherPath -Force
        }
    }

    if (-not (Test-Path -LiteralPath $LauncherPath -PathType Leaf)) {
        throw 'لم يتم العثور على مشغّل DuckAI في المسار الدائم.'
    }

    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'
    if (-not (Test-Path -LiteralPath $wscript -PathType Leaf)) {
        throw 'تعذر العثور على Windows Script Host.'
    }

    $action = New-ScheduledTaskAction -Execute $wscript -Argument ('"' + $LauncherPath + '"')
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'تشغيل مشغل DuckAI المحلي بصمت عند تسجيل الدخول، مع استعادة ذاتية للمكونات عند الحاجة.' -Force | Out-Null

    $registered = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    if ($registered.State -eq 'Disabled') {
        throw 'تم تسجيل المهمة لكنها بقيت معطلة.'
    }

    $actions = @($registered.Actions)
    if ($actions.Count -eq 0) {
        throw 'لم تُسجّل أي مهمة تشغيل فعلية.'
    }

    $registeredAction = $actions[0]
    if ([string]$registeredAction.Execute -notmatch '(?i)wscript\.exe$' -or [string]$registeredAction.Arguments -notmatch 'ResearchTools-Launcher\.vbs') {
        throw 'فشل التحقق من مسار المشغّل المسجل في المهمة.'
    }

    Write-Log 'تم تسجيل مهمة التشغيل التلقائي والتحقق من مسار المشغّل بنجاح.'
}

function Install-LatestRuntime {
    $headers = @{
        'User-Agent' = 'ResearchTools-DuckAI-Bootstrap/3.0'
        'Accept' = 'application/vnd.github+json'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Log 'بدء فحص إصدار DuckAI المنشور على GitHub.'

    $releases = @(Invoke-RestMethod -TimeoutSec 60 -Uri $ReleaseApi -Headers $headers -Method Get)
    $release = $releases |
        Where-Object {
            -not $_.draft -and -not $_.prerelease -and
            [string]$_.tag_name -match '^duckai-runtime-v\d+\.\d+\.\d+$'
        } |
        Sort-Object { [DateTime]::Parse($_.published_at).ToUniversalTime() } -Descending |
        Select-Object -First 1

    if (-not $release) {
        throw 'لم يُعثر على إصدار DuckAI مستقر منشور.'
    }

    $asset = @($release.assets | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1)
    if ($asset.Count -eq 0) {
        throw "لم يُعثر على ملف $AssetName في أحدث إصدار."
    }
    $asset = $asset[0]

    if ([string]$asset.browser_download_url -notmatch '^https://github\.com/moh1983taher-ai/Word-AI-Assistant/releases/download/') {
        throw 'رابط تنزيل الحزمة لا ينتمي إلى مستودع أدوات الباحث المتوقع.'
    }
    if ([int64]$asset.size -gt 100MB -or [int64]$asset.size -lt 1MB) {
        throw 'حجم حزمة DuckAI المنشورة غير متوقع.'
    }
    if ([string]$asset.digest -notmatch '^sha256:([A-Fa-f0-9]{64})$') {
        throw 'لم يوفر GitHub بصمة SHA-256 موثوقة للحزمة؛ أُوقفت الاستعادة.'
    }

    $expectedArchiveHash = $matches[1].ToUpperInvariant()
    $downloadPath = Join-Path $env:TEMP ('DuckAI-Runtime-' + [Guid]::NewGuid().ToString('N') + '.zip')
    $stagePath = Join-Path $RootDir ('DuckAI.Stage-' + [Guid]::NewGuid().ToString('N'))
    $backupPath = Join-Path $RootDir ('DuckAI.Backup-' + [Guid]::NewGuid().ToString('N'))

    try {
        Write-Log ("تنزيل الحزمة: " + $release.tag_name)
        Invoke-WebRequest -UseBasicParsing -TimeoutSec 180 -Uri $asset.browser_download_url -Headers $headers -OutFile $downloadPath

        if ((Get-Item -LiteralPath $downloadPath).Length -ne [int64]$asset.size) {
            throw 'حجم الملف المنزّل لا يطابق الحجم المعلن في GitHub.'
        }

        $actualArchiveHash = (Get-FileHash -LiteralPath $downloadPath -Algorithm SHA256).Hash.ToUpperInvariant()
        if ($actualArchiveHash -ne $expectedArchiveHash) {
            throw 'فشل التحقق من بصمة SHA-256 لحزمة DuckAI؛ لن يتم تثبيتها.'
        }

        if (-not (Test-ZipPaths -ArchivePath $downloadPath)) {
            throw 'تحتوي الحزمة على مسارات غير آمنة.'
        }

        New-Item -ItemType Directory -Path $stagePath -Force | Out-Null
        Expand-Archive -LiteralPath $downloadPath -DestinationPath $stagePath -Force

        if (-not (Test-RuntimeIntegrity -Directory $stagePath)) {
            throw 'فشل التحقق من بصمات مكونات الحزمة الداخلية.'
        }

        $running = @(Get-Process -Name 'duck2api-main', 'ResearchTools-DuckAgent' -ErrorAction SilentlyContinue)
        if ($running.Count -gt 0) {
            throw 'تأجلت الاستعادة لأن إحدى خدمات DuckAI تعمل حالياً؛ أغلق Word وانتظر توقف الخدمات ثم أعد المحاولة.'
        }

        if (Test-Path -LiteralPath $RuntimeDir -PathType Container) {
            Move-Item -LiteralPath $RuntimeDir -Destination $backupPath
        }

        try {
            Move-Item -LiteralPath $stagePath -Destination $RuntimeDir
        }
        catch {
            if (Test-Path -LiteralPath $backupPath -PathType Container) {
                Move-Item -LiteralPath $backupPath -Destination $RuntimeDir
            }
            throw
        }

        if (-not (Test-RuntimeIntegrity -Directory $RuntimeDir)) {
            throw 'لم يجتز مجلد DuckAI الجديد فحص السلامة بعد نقله إلى موضعه النهائي.'
        }

        if (Test-Path -LiteralPath $backupPath -PathType Container) {
            Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction SilentlyContinue
        }

        Write-Log ("اكتملت استعادة DuckAI والتحقق من مكوناته. الإصدار: " + $release.tag_name)
    }
    finally {
        Remove-Item -LiteralPath $downloadPath -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $stagePath -PathType Container) {
            Remove-Item -LiteralPath $stagePath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

$mutex = $null
$hasLock = $false
try {
    New-Item -ItemType Directory -Path $RootDir -Force | Out-Null
    Write-Log 'بدأ DuckAI Bootstrap.'

    $mutex = New-Object System.Threading.Mutex($false, 'Local\ResearchToolsDuckAIBootstrap')
    try {
        $hasLock = $mutex.WaitOne(0)
    }
    catch {
        $hasLock = $false
    }

    if (-not $hasLock) {
        if ($RegisterTask) {
            throw 'توجد عملية استعادة أخرى؛ تعذر تنفيذ التهيئة بأمان.'
        }
        Write-Log 'توجد عملية فحص أو استعادة أخرى؛ انتهت هذه النسخة دون تكرار العمل.'
        exit 0
    }

    if ($RegisterTask) {
        Register-StartupTask
    }

    if ($EnsureRuntime -or -not $RegisterTask) {
        if (Test-RuntimeIntegrity -Directory $RuntimeDir) {
            Write-Log 'المكونات المحلية سليمة؛ لا حاجة إلى التنزيل.'
            exit 0
        }

        Install-LatestRuntime
        if (-not (Test-RuntimeIntegrity -Directory $RuntimeDir)) {
            throw 'انتهت عملية الاستعادة لكن فحص الملفات النهائي لم ينجح.'
        }
        Write-Log 'اجتاز DuckAI فحص السلامة النهائي.'
    }

    exit 0
}
catch {
    Write-Log ('تعذر تجهيز DuckAI: ' + $_.Exception.Message)
    exit 1
}
finally {
    if ($mutex -and $hasLock) {
        try { $mutex.ReleaseMutex() } catch { }
    }
    if ($mutex) {
        $mutex.Dispose()
    }
}
