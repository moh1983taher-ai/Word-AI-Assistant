[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RuntimeArchive,
    [Parameter(Mandatory = $true)][string]$RuntimeShaFile,
    [Parameter(Mandatory = $true)][string]$McpHostArchive,
    [Parameter(Mandatory = $true)][string]$McpHostShaFile,
    [Parameter(Mandatory = $true)][string]$LauncherFile,
    [Parameter(Mandatory = $true)][string]$ManifestFile
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Root = Join-Path $env:LOCALAPPDATA 'ResearchTools'
$Runtime = Join-Path $Root 'DuckAI'
$McpHost = Join-Path $Root 'McpHost'
$Backup = Join-Path $Root 'Backup'
$Logs = Join-Path $Root 'Logs'
$Launcher = Join-Path $Root 'ResearchTools-Launcher.vbs'
$Bootstrap = Join-Path $Root 'DuckAI-Bootstrap.ps1'
$Log = Join-Path $Logs 'ResearchTools-Setup.log'
$TaskName = 'ResearchTools DuckAI Agent'
$AppId = 'f8a582b1-eaf2-44f9-8166-d034255586ad'
$PreferredCatalog = '7A4F5D92-3C81-4B6E-9D27-51C8A6E4F103'
$script:Form = $null
$script:Status = $null
$script:Bar = $null

New-Item -ItemType Directory -Path $Root,$Backup,$Logs -Force | Out-Null

function Log([string]$Message) {
    Add-Content -LiteralPath $Log -Encoding UTF8 -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Message)
}
function Stage([int]$Value,[string]$Message) {
    Log $Message
    if ($script:Status) {
        $script:Status.Text = $Message
        $script:Bar.Value = [Math]::Max(0,[Math]::Min(100,$Value))
        $script:Form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
}
function Show-Progress {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'إعداد أدوات الباحث'
    $f.Width = 600
    $f.Height = 205
    $f.StartPosition = 'CenterScreen'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.ControlBox = $false
    $f.TopMost = $true
    $f.RightToLeft = 'Yes'
    $f.RightToLeftLayout = $true
    $t = New-Object System.Windows.Forms.Label
    $t.Text = 'تثبيت أدوات الباحث والخادم المحلي'
    $t.Font = New-Object System.Drawing.Font('Segoe UI',13,[System.Drawing.FontStyle]::Bold)
    $t.TextAlign = 'MiddleRight'
    $t.SetBounds(24,18,535,32)
    $f.Controls.Add($t)
    $s = New-Object System.Windows.Forms.Label
    $s.Text = 'تهيئة بيئة التثبيت...'
    $s.TextAlign = 'MiddleRight'
    $s.SetBounds(24,65,535,28)
    $f.Controls.Add($s)
    $p = New-Object System.Windows.Forms.ProgressBar
    $p.Minimum = 0
    $p.Maximum = 100
    $p.Value = 2
    $p.SetBounds(24,108,535,20)
    $f.Controls.Add($p)
    $note = New-Object System.Windows.Forms.Label
    $note.Text = 'يرجى عدم إغلاق نافذة التثبيت حتى اكتمال العملية.'
    $note.TextAlign = 'MiddleRight'
    $note.ForeColor = [System.Drawing.Color]::DimGray
    $note.SetBounds(24,137,535,22)
    $f.Controls.Add($note)
    $script:Form = $f
    $script:Status = $s
    $script:Bar = $p
    $f.Show()
    $f.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}
function Test-FileHashes([string]$Directory) {
    $sum = Join-Path $Directory 'SHA256SUMS.txt'
    if (-not (Test-Path -LiteralPath $sum -PathType Leaf)) { return $false }
    $required = @('BUILD-INFO.txt','ResearchTools-DuckAgent.exe','ResearchTools-Launcher.vbs','DuckAI-Bootstrap.ps1','DuckAI-MCP-Bridge.ps1','duck2api-main.exe','manifest.xml','uninstall.ps1')
    $entries = @{}
    foreach ($line in (Get-Content -LiteralPath $sum -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -notmatch '^\s*([A-Fa-f0-9]{64})\s+(.+?)\s*$') { return $false }
        $name = $matches[2].Trim()
        if ([IO.Path]::GetFileName($name) -ne $name -or $name -match '[\\/:]' -or $name -in @('.','..')) { return $false }
        if ($entries.ContainsKey($name)) { return $false }
        $entries[$name] = $matches[1].ToUpperInvariant()
    }
    foreach ($name in $required) { if (-not $entries.ContainsKey($name)) { return $false } }
    foreach ($name in $entries.Keys) {
        $path = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $false }
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToUpperInvariant() -ne $entries[$name]) { return $false }
    }
    return $true
}
function Test-Archive([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        if ($zip.Entries.Count -lt 7 -or $zip.Entries.Count -gt 32) { throw 'عدد الملفات داخل حزمة DuckAI غير متوقع.' }
        [int64]$expanded = 0
        foreach ($entry in $zip.Entries) {
            $name = $entry.FullName
            if ([string]::IsNullOrWhiteSpace($name) -or $name.Contains('/') -or $name.Contains('\') -or $name.Contains(':') -or $entry.Name -ne $name) {
                throw ('مسار غير آمن داخل الحزمة: ' + $name)
            }
            $expanded += [int64]$entry.Length
            if ($expanded -gt 250MB) { throw 'الحجم بعد فك الضغط يتجاوز الحد الآمن.' }
        }
    }
    finally { $zip.Dispose() }
}
function Test-McpHostArchive([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $requiredMcpFiles = @(
            'ResearchTools.McpHost.exe',
            'ResearchTools.McpHost.exe.sha256',
            'Test-McpHost-Preview.ps1',
            'Start-McpHost-Preview.ps1'
        )
        if ($zip.Entries.Count -ne $requiredMcpFiles.Count) {
            throw ('Unexpected MCP archive entry count. Expected ' + $requiredMcpFiles.Count + ', found ' + $zip.Entries.Count + '.')
        }
        $names = @{}
        [int64]$expanded = 0
        foreach ($entry in $zip.Entries) {
            $name = $entry.FullName
            if ([string]::IsNullOrWhiteSpace($name) -or $name.Contains('/') -or $name.Contains('\') -or $name.Contains(':') -or $entry.Name -ne $name) { throw ('Unsafe MCP archive path: ' + $name) }
            if ($names.ContainsKey($name)) { throw ('Duplicate MCP archive entry: ' + $name) }
            $names[$name] = $true
            $expanded += [int64]$entry.Length
            if ($expanded -gt 180MB) { throw 'MCP archive expands beyond the safe size limit.' }
        }
        foreach ($requiredMcpFile in $requiredMcpFiles) {
            if (-not $names.ContainsKey($requiredMcpFile)) {
                throw ('MCP archive is missing required file: ' + $requiredMcpFile)
            }
        }
    }
    finally { $zip.Dispose() }
}
function Install-McpHost {
    if (-not (Test-Path -LiteralPath $McpHostArchive -PathType Leaf) -or -not (Test-Path -LiteralPath $McpHostShaFile -PathType Leaf)) { throw 'MCP host package files are missing.' }
    $sidecar = Get-Content -LiteralPath $McpHostShaFile -Raw
    if ($sidecar -notmatch '(?i)^\s*([a-f0-9]{64})\s+ResearchTools-MCP-Host-release\.zip\s*$') { throw 'Invalid MCP host archive checksum file.' }
    $expectedArchiveHash = $matches[1].ToUpperInvariant()
    if ((Get-FileHash -LiteralPath $McpHostArchive -Algorithm SHA256).Hash.ToUpperInvariant() -ne $expectedArchiveHash) { throw 'MCP host archive SHA-256 mismatch.' }
    Test-McpHostArchive -Path $McpHostArchive
    $stage = Join-Path $Root ('McpHost.Stage-' + [Guid]::NewGuid().ToString('N'))
    $previous = Join-Path $Backup ('McpHost-Previous-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        [IO.Compression.ZipFile]::ExtractToDirectory($McpHostArchive,$stage)
        $exe = Join-Path $stage 'ResearchTools.McpHost.exe'
        $internalSha = Join-Path $stage 'ResearchTools.McpHost.exe.sha256'
        if (-not (Test-Path -LiteralPath $exe -PathType Leaf) -or -not (Test-Path -LiteralPath $internalSha -PathType Leaf)) { throw 'MCP host files are incomplete after extraction.' }
        $inner = Get-Content -LiteralPath $internalSha -Raw
        if ($inner -notmatch '(?i)^\s*([a-f0-9]{64})\s+ResearchTools\.McpHost\.exe\s*$') { throw 'Invalid MCP host executable checksum.' }
        $expectedExeHash = $matches[1].ToUpperInvariant()
        if ((Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash.ToUpperInvariant() -ne $expectedExeHash) { throw 'MCP host executable SHA-256 mismatch.' }
        if (Test-Path -LiteralPath $McpHost -PathType Container) { Move-Item -LiteralPath $McpHost -Destination $previous }
        try {
            Move-Item -LiteralPath $stage -Destination $McpHost
            $installedExe = Join-Path $McpHost 'ResearchTools.McpHost.exe'
            if ((Get-FileHash -LiteralPath $installedExe -Algorithm SHA256).Hash.ToUpperInvariant() -ne $expectedExeHash) { throw 'Installed MCP host verification failed.' }
        }
        catch {
            if (Test-Path -LiteralPath $McpHost -PathType Container) { Remove-Item -LiteralPath $McpHost -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $previous -PathType Container) { Move-Item -LiteralPath $previous -Destination $McpHost -Force }
            throw
        }
    }
    finally {
        if (Test-Path -LiteralPath $stage -PathType Container) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Log 'MCP host installed and internal/external hashes verified.'
}
function Install-DurableLauncher {
    if (-not (Test-Path -LiteralPath $LauncherFile -PathType Leaf)) { throw 'Unified ResearchTools launcher is missing.' }
    $content = Get-Content -LiteralPath $LauncherFile -Raw
    $bootstrapHash = (Get-FileHash -LiteralPath $Bootstrap -Algorithm SHA256).Hash.ToLowerInvariant()
    $pinPattern = '(?im)^[ \t]*Const[ \t]+BOOTSTRAP_SHA256[ \t]*=[ \t]*"([a-f0-9]{64})"'
    $pinMatch = [regex]::Match($content,$pinPattern)
    if (-not $pinMatch.Success) { throw 'Unified launcher has no valid BOOTSTRAP_SHA256 declaration.' }
    $declaredHash = $pinMatch.Groups[1].Value.ToLowerInvariant()
    if ($declaredHash -ne $bootstrapHash) {
        # DuckAI archive integrity was verified before this point. Align the
        # launcher pin to that verified bootstrap, then verify the edit.
        $replacement = 'Const BOOTSTRAP_SHA256 = "' + $bootstrapHash + '"'
        $content = [regex]::Replace($content,$pinPattern,$replacement,1)
        $encoding = [System.Text.UTF8Encoding]::new($true)
        [IO.File]::WriteAllText($LauncherFile,$content,$encoding)
        Log ('Unified launcher pin synchronized with verified DuckAI bootstrap. Previous=' + $declaredHash + '; Current=' + $bootstrapHash)
    }
    $verifiedContent = Get-Content -LiteralPath $LauncherFile -Raw
    $verifiedPin = [regex]::Match($verifiedContent,$pinPattern)
    if (-not $verifiedPin.Success -or $verifiedPin.Groups[1].Value.ToLowerInvariant() -ne $bootstrapHash) {
        throw 'Unified launcher bootstrap pin could not be synchronized with the verified DuckAI bootstrap.'
    }
    if ($verifiedContent -notmatch 'ResearchTools\.McpHost\.exe' -or $verifiedContent -notmatch 'RESEARCHTOOLS_MCPHOST_V1') {
        throw 'Launcher does not include the new MCP host integration.'
    }
    Copy-Item -LiteralPath $LauncherFile -Destination $Launcher -Force
    Log 'Unified DuckAI and MCP background launcher installed and bootstrap pin verified.'
}
function Get-PowerShell {
    $p = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ($env:PROCESSOR_ARCHITEW6432) {
        $native = Join-Path $env:SystemRoot 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path -LiteralPath $native) { return $native }
    }
    return $p
}
function Test-Writable([string]$Path) {
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
        $probe = Join-Path $Path ('.rt-' + [Guid]::NewGuid().ToString('N') + '.tmp')
        [IO.File]::WriteAllText($probe,'ok')
        Remove-Item -LiteralPath $probe -Force
        return $true
    }
    catch { return $false }
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
}
else {
    New-SmbShare -Name $ShareName -Path $SharePath -FullAccess $Account -Description 'ResearchTools Office add-in catalog' -ErrorAction Stop | Out-Null
}
'@
    Set-Content -LiteralPath $helper -Value $code -Encoding UTF8
    try {
        $args = '-NoProfile -ExecutionPolicy Bypass -File "' + $helper + '" -ShareName "' + $Name + '" -SharePath "' + $Path + '" -Account "' + $account + '"'
        $p = Start-Process -FilePath (Get-PowerShell) -Verb RunAs -ArgumentList $args -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw ('تعذر إنشاء مجلد المشاركة. رمز الخروج: ' + $p.ExitCode) }
    }
    finally { Remove-Item -LiteralPath $helper -Force -ErrorAction SilentlyContinue }
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
            }
            catch { }
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
    }
    catch { }
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
        }
        catch { }
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
function Stop-OldProcesses {
    foreach ($proc in @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)) {
        $name = [string]$proc.Name
        $path = [string]$proc.ExecutablePath
        $cmd = [string]$proc.CommandLine
        $runtimeProc = ($name -in @('duck2api-main.exe','ResearchTools-DuckAgent.exe')) -and $path -and $path.StartsWith($Runtime,[StringComparison]::OrdinalIgnoreCase)
        $launcherProc = ($name -ieq 'wscript.exe') -and $cmd -and $cmd.IndexOf($Launcher,[StringComparison]::OrdinalIgnoreCase) -ge 0
        $legacyBridgeProc = ($name -ieq 'powershell.exe') -and $cmd -and $cmd.IndexOf('DuckAI-MCP-Bridge.ps1',[StringComparison]::OrdinalIgnoreCase) -ge 0
        $mcpHostProc = ($name -ieq 'ResearchTools.McpHost.exe') -and $path -and $path.StartsWith($McpHost,[StringComparison]::OrdinalIgnoreCase)
        if ($runtimeProc -or $launcherProc -or $legacyBridgeProc -or $mcpHostProc) { Stop-Process -Id $proc.ProcessId -Force -ErrorAction SilentlyContinue }
    }
}
function Install-Runtime {
    if (-not (Test-Path -LiteralPath $RuntimeArchive -PathType Leaf) -or -not (Test-Path -LiteralPath $RuntimeShaFile -PathType Leaf)) { throw 'ملفات الخادم المضمنة غير مكتملة.' }
    $sidecar = Get-Content -LiteralPath $RuntimeShaFile -Raw
    if ($sidecar -notmatch '(?i)^\s*([a-f0-9]{64})\s+DuckAI-Runtime-release\.zip\s*$') { throw 'صيغة ملف SHA-256 غير صحيحة.' }
    $expected = $matches[1].ToUpperInvariant()
    $actual = (Get-FileHash -LiteralPath $RuntimeArchive -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne $expected) { throw 'فشل التحقق من SHA-256 لحزمة DuckAI.' }
    Test-Archive -Path $RuntimeArchive
    $stage = Join-Path $Root ('DuckAI.Stage-' + [Guid]::NewGuid().ToString('N'))
    $previous = Join-Path $Backup ('DuckAI-Previous-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        [IO.Compression.ZipFile]::ExtractToDirectory($RuntimeArchive,$stage)
        if (-not (Test-FileHashes -Directory $stage)) { throw 'فشل التحقق من البصمات الداخلية لملفات الخادم.' }
        if (Test-Path -LiteralPath $Runtime -PathType Container) { Move-Item -LiteralPath $Runtime -Destination $previous }
        try {
            Move-Item -LiteralPath $stage -Destination $Runtime
            if (-not (Test-FileHashes -Directory $Runtime)) { throw 'فشل الفحص النهائي لمجلد الخادم.' }
        }
        catch {
            if (Test-Path -LiteralPath $Runtime) { Remove-Item -LiteralPath $Runtime -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $previous) { Move-Item -LiteralPath $previous -Destination $Runtime -Force }
            throw
        }
    }
    finally {
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $runtimeLauncherFile = Join-Path $Runtime 'ResearchTools-Launcher.vbs'
    if (-not (Test-Path -LiteralPath $runtimeLauncherFile -PathType Leaf)) { throw 'ملف مشغّل DuckAI داخل الحزمة غير موجود.' }
    Copy-Item -LiteralPath (Join-Path $Runtime 'DuckAI-Bootstrap.ps1') -Destination $Bootstrap -Force
    $hash = (Get-FileHash -LiteralPath $Bootstrap -Algorithm SHA256).Hash.ToLowerInvariant()
    if (-not (Get-Content -LiteralPath $runtimeLauncherFile -Raw).Contains($hash)) { throw 'بصمة Bootstrap داخل مشغّل حزمة DuckAI لا تطابق الملف المثبت.' }
    if (-not (Test-FileHashes -Directory $Runtime)) { throw 'فشل التحقق النهائي من ملفات DuckAI.' }
}

try {
    Show-Progress
    Stage 5 'التحقق من متطلبات التثبيت وإغلاق Word...'
    if (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Count -gt 0) { throw 'أغلق Microsoft Word بالكامل ثم أعد تشغيل Setup.' }
    Stage 15 'التحقق من ملفات التثبيت...'
    if (-not (Test-Path -LiteralPath $ManifestFile -PathType Leaf)) { throw 'ملف Manifest غير موجود.' }
    Stage 25 'إيقاف عمليات DuckAI السابقة دون المساس بملفات التطوير...'
    Stop-OldProcesses
    Stage 34 'تثبيت محرك DuckAI المحلي والتحقق من سلامته...'
    Install-Runtime
    Stage 48 'تثبيت مضيف MCP الجديد وتهيئة اتصاله بـ DuckAI...'
    Install-McpHost
    Stage 58 'تثبيت مشغّل الخدمات الموحّد والتحقق من بصمته...'
    Install-DurableLauncher
    Stage 68 'فحص مجلد Word المشترك وتهيئة الكتالوج الموثوق...'
    $catalog = Get-Catalog
    Stage 76 'إضافة أدوات الباحث إلى قائمة الإضافات المشتركة في Word...'
    Install-Manifest -Catalog $catalog
    Stage 86 'تسجيل مهمة التشغيل التلقائي لخادم DuckAI...'
    $powershell = Get-PowerShell
    $arguments = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $Bootstrap + '" -RegisterTask'
    $bootstrapProcess = Start-Process -FilePath $powershell -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru
    if ($bootstrapProcess.ExitCode -ne 0) { throw ('فشل تسجيل مهمة التشغيل التلقائي. ExitCode=' + $bootstrapProcess.ExitCode) }
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    if ($task.State -eq 'Disabled') { throw 'مهمة تشغيل DuckAI مسجلة لكنها معطلة.' }
    $action = @($task.Actions) | Select-Object -First 1
    if (-not $action -or [string]$action.Execute -notmatch '(?i)wscript\.exe$' -or [string]$action.Arguments -notmatch 'ResearchTools-Launcher\.vbs') { throw 'مهمة التشغيل لا تشير إلى المشغّل المثبت.' }
    Stage 94 'تشغيل المشغّل الدائم والتحقق من وجوده...'
    $wscript = Join-Path $env:WINDIR 'System32\wscript.exe'
    Start-Process -FilePath $wscript -ArgumentList ('"' + $Launcher + '"') -WindowStyle Hidden | Out-Null
    Start-Sleep -Seconds 2
    $running = @(Get-CimInstance Win32_Process -Filter "Name='wscript.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($Launcher,[StringComparison]::OrdinalIgnoreCase) -ge 0 })
    if ($running.Count -eq 0) { throw 'تعذر تشغيل المشغّل الدائم.' }
    Stage 100 'اكتمل تثبيت أدوات الباحث وDuckAI وMCP وإظهارها في Word.'
    Log 'SUCCESS: DuckAI runtime, MCP host, durable launcher, Word shared catalog, manifest and startup task installed.'
    Start-Sleep -Milliseconds 1000
    $script:Form.Close()
    exit 0
}
catch {
    $message = $_.Exception.Message
    try { Log ('ERROR: ' + $message) } catch { }
    try {
        $details = 'تعذر إكمال التثبيت:' + [Environment]::NewLine + $message + [Environment]::NewLine + [Environment]::NewLine + 'راجع السجل:' + [Environment]::NewLine + $Log
        [System.Windows.Forms.MessageBox]::Show($details,'تعذر تثبيت أدوات الباحث',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    catch { Write-Error $message }
    if ($script:Form) { $script:Form.Close() }
    exit 1
}
