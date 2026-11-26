[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ManifestFile,
    [Parameter(Mandatory = $true)][string]$ProviderLauncherFile
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Root = Join-Path $env:LOCALAPPDATA 'ResearchTools'
$DuckRuntime = Join-Path $Root 'DuckAI'
$McpHost = Join-Path $Root 'McpHost'
$Backup = Join-Path $Root 'Backup'
$Logs = Join-Path $Root 'Logs'
$Log = Join-Path $Logs 'ResearchTools-Setup.log'
$LegacyTaskName = 'ResearchTools DuckAI Agent'
$ProviderTaskName = 'ResearchTools DuckAI Provider'
$ProviderLauncher = Join-Path $Root 'ResearchTools-DuckAI-Provider.vbs'
$AppId = 'f8a582b1-eaf2-44f9-8166-d034255586ad'
$PreferredCatalog = '7A4F5D92-3C81-4B6E-9D27-51C8A6E4F103'
New-Item -ItemType Directory -Path $Root,$Backup,$Logs -Force | Out-Null

function Log([string]$Message) {
    Add-Content -LiteralPath $Log -Encoding UTF8 -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Message)
}
function Get-PowerShell {
    $path = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ($env:PROCESSOR_ARCHITEW6432) {
        $native = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path -LiteralPath $native -PathType Leaf) { return $native }
    }
    return $path
}
function Test-PathWithin([string]$Candidate,[string]$Directory) {
    if ([string]::IsNullOrWhiteSpace($Candidate) -or [string]::IsNullOrWhiteSpace($Directory)) { return $false }
    try {
        $candidatePath = [IO.Path]::GetFullPath($Candidate)
        $directoryPath = [IO.Path]::GetFullPath($Directory).TrimEnd('\') + '\'
        return $candidatePath.StartsWith($directoryPath,[StringComparison]::OrdinalIgnoreCase)
    } catch { return $false }
}

function Remove-LegacyLocalMcpServices {
    $legacyLaunchers = @(
        (Join-Path $Root 'ResearchTools-Launcher.vbs'),
        (Join-Path $DuckRuntime 'ResearchTools-Launcher.vbs')
    )
    foreach ($taskName in @($LegacyTaskName,$ProviderTaskName)) {
        try { Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue } catch { Log $_.Exception.Message }
        try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch { Log $_.Exception.Message }
    }

    $stopped = 0
    foreach ($process in @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)) {
        $name = [string]$process.Name
        $exe = [string]$process.ExecutablePath
        $cmd = [string]$process.CommandLine
        $legacyAgent = ($name -ieq 'ResearchTools-DuckAgent.exe') -and (Test-PathWithin $exe $DuckRuntime)
        $mcpHostProcess = ($name -ieq 'ResearchTools.McpHost.exe') -and (Test-PathWithin $exe $McpHost)
        $bridge = ($name -ieq 'powershell.exe') -and
            $cmd.IndexOf('DuckAI-MCP-Bridge.ps1',[StringComparison]::OrdinalIgnoreCase) -ge 0 -and
            $cmd.IndexOf($Root,[StringComparison]::OrdinalIgnoreCase) -ge 0
        $legacyLauncher = $false
        if ($name -ieq 'wscript.exe') {
            foreach ($legacyLauncherPath in $legacyLaunchers) {
                if ($cmd.IndexOf($legacyLauncherPath,[StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $legacyLauncher = $true
                    break
                }
            }
        }
        if ($legacyAgent -or $mcpHostProcess -or $bridge -or $legacyLauncher) {
            try {
                Stop-Process -Id ([int]$process.ProcessId) -Force -ErrorAction Stop
                $stopped++
            } catch { Log ('تعذر إيقاف عملية MCP محلية PID=' + $process.ProcessId + ': ' + $_.Exception.Message) }
        }
    }

    Start-Sleep -Milliseconds 500
    if (Test-Path -LiteralPath $McpHost -PathType Container) {
        Remove-Item -LiteralPath $McpHost -Recurse -Force
        if (Test-Path -LiteralPath $McpHost -PathType Container) { throw ('تعذر إزالة مجلد مضيف MCP المحلي: ' + $McpHost) }
    }

    $runtimeMcpFiles = @(
        'ResearchTools-DuckAgent.exe',
        'ResearchTools-Launcher.vbs',
        'DuckAI-MCP-Bridge.ps1',
        'DuckAI-MCP-Bridge.log',
        'DuckAI-Bootstrap.ps1',
        'Apply-DuckAI-AutoRepair.ps1',
        'Install-DuckAI-AutoRepair.vbs',
        'DuckAI-Update.vbs'
    )
    foreach ($name in $runtimeMcpFiles) {
        $path = Join-Path $DuckRuntime $name
        if (Test-Path -LiteralPath $path -PathType Leaf) { Remove-Item -LiteralPath $path -Force }
    }

    foreach ($name in @(
        'ResearchTools-Launcher.vbs','DuckAI-Bootstrap.ps1','DuckAI-Bootstrap.log',
        'DuckAI-Installed-Release.txt','DuckAI-Update.vbs','DuckAI-Update.log',
        'DuckAI-Update-Result.txt','DuckAI-MCP-Bridge.log','DuckAI-MCP-Bridge.pid',
        'DuckAI-AutoRepair-Install.log','DuckAI-AutoRepair-PowerShell.log'
    )) {
        $path = Join-Path $Root $name
        if (Test-Path -LiteralPath $path -PathType Leaf) { Remove-Item -LiteralPath $path -Force }
    }

    $oldRootBackup = Join-Path $Root 'DuckAI-old-backup'
    if (Test-Path -LiteralPath $oldRootBackup -PathType Container) { Remove-Item -LiteralPath $oldRootBackup -Recurse -Force }
    foreach ($pattern in @('DuckAI-Previous-*','McpHost-Previous-*')) {
        foreach ($old in @(Get-ChildItem -LiteralPath $Backup -Directory -Filter $pattern -ErrorAction SilentlyContinue)) {
            Remove-Item -LiteralPath $old.FullName -Recurse -Force
        }
    }
    Log ('أزيل مضيف MCP وجسره والوكيل القديم، وحُفظ محرك Duck2API العام. العمليات الموقوفة: ' + $stopped)
}

function Install-DuckAIProvider {
    $duckExe = Join-Path $DuckRuntime 'duck2api-main.exe'
    if (-not (Test-Path -LiteralPath $ProviderLauncherFile -PathType Leaf)) { throw 'ملف تشغيل مزود Duck.ai العام غير موجود.' }
    if (-not (Test-Path -LiteralPath $duckExe -PathType Leaf)) {
        if (Test-Path -LiteralPath $ProviderLauncher -PathType Leaf) { Remove-Item -LiteralPath $ProviderLauncher -Force }
        Log 'لم يعثر المثبت على duck2api-main.exe؛ لم يسجل مهمة تشغيل Duck.ai. تظل مزودات السحابة متاحة.'
        return
    }

    Copy-Item -LiteralPath $ProviderLauncherFile -Destination $ProviderLauncher -Force
    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $action = New-ScheduledTaskAction -Execute $wscript -Argument ('"' + $ProviderLauncher + '"')
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $currentUser
    $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::FromMinutes(2))
    Register-ScheduledTask -TaskName $ProviderTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'Starts the standalone Duck.ai model endpoint only; it does not start local MCP services.' -Force | Out-Null
    Start-ScheduledTask -TaskName $ProviderTaskName
    Log 'سجل المثبت مهمة واحدة لمزود Duck.ai العام فقط؛ لا تشغل مضيف MCP أو جسره.'
}

function Test-Writable([string]$Path) {
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
        $probe = Join-Path $Path ('.rt-' + [Guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText($probe,'ok')
        Remove-Item -LiteralPath $probe -Force
        return $true
    } catch { return $false }
}
function Create-Share([string]$Name,[string]$Path) {
    $account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $helper = Join-Path $env:TEMP ('ResearchTools-Share-' + [Guid]::NewGuid().ToString('N') + '.ps1')
    $code = @'
param([string]$ShareName,[string]$SharePath,[string]$Account)
$ErrorActionPreference = 'Stop'
Import-Module SmbShare -ErrorAction Stop
$existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($existing) {
 if (-not ([IO.Path]::GetFullPath([string]$existing.Path).TrimEnd('\')).Equals([IO.Path]::GetFullPath($SharePath).TrimEnd('\'),[StringComparison]::OrdinalIgnoreCase)) { throw 'مشاركة ResearchTools موجودة لمسار آخر.' }
} else { New-SmbShare -Name $ShareName -Path $SharePath -FullAccess $Account -Description 'ResearchTools Office add-in catalog' -ErrorAction Stop | Out-Null }
'@
    Set-Content -LiteralPath $helper -Value $code -Encoding UTF8
    try {
        $args = '-NoProfile -ExecutionPolicy Bypass -File "' + $helper + '" -ShareName "' + $Name + '" -SharePath "' + $Path + '" -Account "' + $account + '"'
        $p = Start-Process -FilePath (Get-PowerShell) -Verb RunAs -ArgumentList $args -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw ('تعذر إنشاء مجلد المشاركة. ExitCode=' + $p.ExitCode) }
    } finally { Remove-Item -LiteralPath $helper -Force -ErrorAction SilentlyContinue }
}
function Get-Catalog {
    $registry = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'
    $found = @()
    if (Test-Path -LiteralPath $registry) {
        foreach ($key in @(Get-ChildItem -LiteralPath $registry -ErrorAction SilentlyContinue)) {
            try {
                $props = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction Stop
                $url = [string]$props.Url
                if ($url.StartsWith('\\') -and (Test-Writable $url)) {
                    $id = [string]$props.Id
                    if (-not $id) { $id = $key.PSChildName }
                    $found += [pscustomobject]@{Path=$url;Id=$id;Key=$key.PSPath}
                }
            } catch { }
        }
    }
    $chosen = $found | Sort-Object @{Expression={if (([string]$_.Id).Replace('{','').Replace('}','') -ieq $PreferredCatalog) {0} else {1}}} | Select-Object -First 1
    if ($chosen) {
        New-ItemProperty -LiteralPath $chosen.Key -Name Flags -PropertyType DWord -Value 1 -Force | Out-Null
        Log ('إعادة استخدام كتالوج Word المشترك: ' + $chosen.Path)
        return [string]$chosen.Path
    }
    $local = Join-Path $Root 'OfficeAddins'
    New-Item -ItemType Directory -Path $local -Force | Out-Null
    $shareName = 'ResearchToolsAddins'
    try {
        Import-Module SmbShare -ErrorAction Stop
        $share = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
        if ($share -and -not ([IO.Path]::GetFullPath([string]$share.Path).TrimEnd('\')).Equals([IO.Path]::GetFullPath($local).TrimEnd('\'),[StringComparison]::OrdinalIgnoreCase)) {
            $suffix = ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -replace '[^A-Fa-f0-9]','')
            $shareName = 'ResearchToolsAddins-' + $suffix.Substring([Math]::Max(0,$suffix.Length - 6))
        }
    } catch { }
    Create-Share -Name $shareName -Path $local
    $unc = '\\' + $env:COMPUTERNAME + '\' + $shareName
    if (-not (Test-Writable $unc)) { throw ('تعذر الوصول إلى مجلد المشاركة بعد إنشائه: ' + $unc) }
    New-Item -Path $registry -Force | Out-Null
    $newKey = Join-Path $registry '{4B78E8D1-1F8A-43B0-9859-4C10A5E127D2}'
    New-Item -Path $newKey -Force | Out-Null
    New-ItemProperty -LiteralPath $newKey -Name Id -PropertyType String -Value '{4B78E8D1-1F8A-43B0-9859-4C10A5E127D2}' -Force | Out-Null
    New-ItemProperty -LiteralPath $newKey -Name Url -PropertyType String -Value $unc -Force | Out-Null
    New-ItemProperty -LiteralPath $newKey -Name Flags -PropertyType DWord -Value 1 -Force | Out-Null
    Log ('إنشاء كتالوج Word مشترك جديد: ' + $unc)
    return $unc
}
function Install-Manifest([string]$Catalog) {
    [xml]$xml = Get-Content -LiteralPath $ManifestFile -Raw
    $id = $xml.SelectSingleNode("//*[local-name()='Id']")
    $source = $xml.SelectSingleNode("//*[local-name()='DefaultSettings']/*[local-name()='SourceLocation']")
    if (-not $id -or $id.InnerText -ne $AppId) { throw 'معرّف إضافة Word غير مطابق.' }
    if (-not $source -or $source.GetAttribute('DefaultValue') -notmatch '^https://moh1983taher-ai\.github\.io/Word-AI-Assistant/taskpane\.html$') { throw 'رابط تشغيل إضافة Word غير صحيح.' }
    $filename = 'Word-AI-Assistant.xml'
    foreach ($file in @(Get-ChildItem -LiteralPath $Catalog -Filter '*.xml' -File -ErrorAction Stop)) {
        try {
            [xml]$existing = Get-Content -LiteralPath $file.FullName -Raw
            $existingId = $existing.SelectSingleNode("//*[local-name()='Id']")
            if ($existingId -and $existingId.InnerText -eq $AppId) { $filename = $file.Name; break }
        } catch { }
    }
    $destination = Join-Path $Catalog $filename
    if (Test-Path -LiteralPath $destination -PathType Leaf) {
        Copy-Item -LiteralPath $destination -Destination (Join-Path $Backup ($filename + '.' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.bak')) -Force
    }
    $temp = $destination + '.new'
    Copy-Item -LiteralPath $ManifestFile -Destination $temp -Force
    Move-Item -LiteralPath $temp -Destination $destination -Force
    Log ('تثبيت Manifest في مجلد Word المشترك: ' + $destination)
}

try {
    if (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -gt 0) { throw 'أغلق Microsoft Word بالكامل ثم أعد تشغيل Setup.' }
    if (-not (Test-Path -LiteralPath $ManifestFile -PathType Leaf)) { throw 'ملف Manifest غير موجود.' }
    if (-not (Test-Path -LiteralPath $ProviderLauncherFile -PathType Leaf)) { throw 'ملف تشغيل Duck.ai العام غير موجود.' }
    Log 'بدء تثبيت إضافة Word وتنظيف مكونات MCP المحلية مع الحفاظ على مزود Duck.ai العام.'
    Remove-LegacyLocalMcpServices
    Install-DuckAIProvider
    $catalog = Get-Catalog
    Install-Manifest -Catalog $catalog
    Log 'اكتمل التثبيت. البحث في المصادر عبر خادمي MCP السحابيين؛ مزود Duck.ai العام مستقل عن مسار البحث.'
    exit 0
}
catch {
    $message = $_.Exception.Message
    try { Log ('ERROR: ' + $message) } catch { }
    Write-Error $message
    exit 1
}
