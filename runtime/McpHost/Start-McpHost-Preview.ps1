[CmdletBinding()]
param(
    [string]$Question = 'ابحث في الجامع عن كتب تتناول مقاصد الشريعة، واذكر أسماء الكتب ومؤلفيها مع توثيق ما يتاح من بيانات.',
    [string]$Model = '',
    [string[]]$Servers = @('aljam3', 'shamela')
)

$ErrorActionPreference = 'Stop'
$port = 18082
$base = "http://127.0.0.1:$port"
$hostPath = Join-Path $PSScriptRoot 'ResearchTools.McpHost.exe'
$testPath = Join-Path $PSScriptRoot 'Test-McpHost-Preview.ps1'

if (-not (Test-Path -LiteralPath $hostPath -PathType Leaf)) {
    throw "لم يُعثر على المضيف الجديد: $hostPath"
}
if (-not (Test-Path -LiteralPath $testPath -PathType Leaf)) {
    throw "لم يُعثر على سكربت الاختبار: $testPath"
}

$existingListener = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
if ($existingListener) {
    throw "المنفذ $port مستخدم بالفعل. أوقف التطبيق الذي يستخدمه ثم أعد الاختبار؛ لم يُنهِ هذا السكربت أي عملية."
}

$previousPort = $env:RESEARCHTOOLS_MCP_PORT
$previousPreviewLogId = $env:RESEARCHTOOLS_MCP_PREVIEW_LOG_ID
$previewLogId = [Guid]::NewGuid().ToString('N')
$previewLog = Join-Path $env:LOCALAPPDATA "ResearchTools\DuckAI\DuckAI-MCP-Host-preview-$previewLogId.log"
$hostProcess = $null

try {
    $env:RESEARCHTOOLS_MCP_PORT = [string]$port
    $env:RESEARCHTOOLS_MCP_PREVIEW_LOG_ID = $previewLogId
    $hostProcess = Start-Process -FilePath $hostPath -WorkingDirectory $PSScriptRoot -PassThru -WindowStyle Hidden
}
finally {
    if ($null -eq $previousPort) {
        Remove-Item Env:RESEARCHTOOLS_MCP_PORT -ErrorAction SilentlyContinue
    }
    else {
        $env:RESEARCHTOOLS_MCP_PORT = $previousPort
    }

    if ($null -eq $previousPreviewLogId) {
        Remove-Item Env:RESEARCHTOOLS_MCP_PREVIEW_LOG_ID -ErrorAction SilentlyContinue
    }
    else {
        $env:RESEARCHTOOLS_MCP_PREVIEW_LOG_ID = $previousPreviewLogId
    }
}

try {
    Write-Host 'تشغيل المضيف التجريبي على 127.0.0.1:18082 دون المساس بالخدمة الحالية...'
    $ready = $false

    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        if ($hostProcess.HasExited) {
            throw "توقف المضيف التجريبي عند بدء التشغيل. راجع سجل DuckAI-MCP-Host.log."
        }

        try {
            $ping = Invoke-RestMethod -Uri ($base + '/ping') -TimeoutSec 2
            if ($ping.listening -eq $true -and $ping.service -eq 'ResearchTools.McpHost') {
                $ready = $true
                break
            }
        }
        catch {
            Start-Sleep -Seconds 1
        }
    }

    if (-not $ready) {
        throw 'لم يستجب المضيف التجريبي خلال المهلة المحددة.'
    }

    Write-Host 'بدأ المضيف بنجاح. تنفيذ اختبارات DuckAI وMCP والبحث وتزامن /ping...'
    Write-Host ('مصادر الاختبار: ' + ($Servers -join ' + '))
    Write-Host ('معرّف سجل الاختبار المعزول: ' + $previewLogId)
    & $testPath -Question $Question -Model $Model -Servers $Servers
}
catch {
    if (Test-Path -LiteralPath $previewLog -PathType Leaf) {
        Write-Host ([Environment]::NewLine + '========== سجل المضيف التجريبي المعزول ==========')
        Get-Content -LiteralPath $previewLog -Tail 120 | ForEach-Object { Write-Host $_ }
    }
    throw
}
finally {
    if ($hostProcess -and -not $hostProcess.HasExited) {
        Stop-Process -Id $hostProcess.Id -Force -ErrorAction SilentlyContinue
        Wait-Process -Id $hostProcess.Id -Timeout 5 -ErrorAction SilentlyContinue
    }

    Write-Host 'أُغلقت عملية المضيف التجريبية فقط. خدمة أدوات الباحث الحالية لم تتغير.'
}
