[CmdletBinding()]
param(
    [string]$Question = 'ابحث في الجامع عن كتب تتناول مقاصد الشريعة، واذكر أسماء الكتب ومؤلفيها مع توثيق ما يتاح من بيانات.',
    [string]$Model = '',
    [string[]]$Servers = @('aljam3', 'shamela')
)

$ErrorActionPreference = 'Stop'
$base = 'http://127.0.0.1:18082'
$origin = 'https://moh1983taher-ai.github.io'
$headers = @{ Origin = $origin }

Write-Host '1/4 - اختبار حياة المضيف الجديد...'
$ping = Invoke-RestMethod -Uri ($base + '/ping') -TimeoutSec 5
if ($ping.listening -ne $true) { throw 'MCP host did not report listening=true.' }
Write-Host "OK: $($ping.service) v$($ping.version)"

Write-Host ([Environment]::NewLine + '2/4 - اختبار DuckAI المحلي...')
$health = Invoke-RestMethod -Uri ($base + '/health') -TimeoutSec 15
if ($health.ready -ne $true) { throw "DuckAI is not ready: $($health.error)" }
Write-Host "OK: provider=$($health.provider); model=$($health.model)"

Write-Host ([Environment]::NewLine + '3/4 - اكتشاف أدوات MCP من المضيف نفسه...')
$sources = Invoke-RestMethod -Uri ($base + '/v1/mcp/sources') -TimeoutSec 120
$sources.sources | Format-Table id, label, connected, toolCount -AutoSize
if ($sources.connected -ne 2 -or $sources.totalTools -lt 11) {
    throw 'Both MCP sources and their read-only catalogs were not discovered.'
}

Write-Host ([Environment]::NewLine + '4/4 - تنفيذ بحث حقيقي عبر DuckAI وMCP...')
Write-Host ('المكتبات المختارة: ' + ($Servers -join ' + '))
$payload = @{
    client = 'word-addin'
    servers = @($Servers)
    model = $Model
    messages = @(
        @{ role = 'user'; text = $Question }
    )
}
$body = ConvertTo-Json -InputObject $payload -Depth 40 -Compress
$job = Start-Job -ArgumentList $base, $headers, $body -ScriptBlock {
    param($uri, $requestHeaders, $requestBody)
    Invoke-RestMethod -Method Post -Uri ($uri + '/v1/research') -Headers $requestHeaders -ContentType 'text/plain; charset=utf-8' -Body $requestBody -TimeoutSec 190
}

$pingCount = 0
try {
    while ($job.State -eq 'Running') {
        $live = Invoke-RestMethod -Uri ($base + '/ping') -TimeoutSec 3
        if ($live.listening -ne $true) { throw 'Liveness check failed during active research.' }
        $pingCount++
        Start-Sleep -Milliseconds 800
    }

    $result = Receive-Job -Job $job -Wait
}
finally {
    if ($job) {
        if ($job.State -eq 'Running') {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
        }
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

if ([string]::IsNullOrWhiteSpace([string]$result.answer)) {
    throw 'Research completed without a non-empty answer.'
}

Write-Host ([Environment]::NewLine + '========== نتيجة الاختبار ==========')
Write-Host "نجح البحث عبر النموذج: $($result.model)"
Write-Host "عدد فحوص /ping أثناء البحث: $pingCount"
Write-Host ([Environment]::NewLine + 'الأدوات التي استُدعيت:')
if (@($result.toolCalls).Count -eq 0) {
    throw 'أعاد DuckAI إجابة دون تنفيذ أداة MCP بنجاح؛ فشل اختبار البحث الفعلي، ولا تُعد هذه الإجابة بحثاً موثقاً.'
}
$result.toolCalls | Format-Table server, label, tool -AutoSize
$unexpectedServers = @($result.toolCalls | Where-Object { $Servers -notcontains $_.server })
if ($unexpectedServers.Count -gt 0) {
    throw 'The host invoked a source that the user did not select.'
}
Write-Host 'تم احترام اختيار المستخدم للمكتبات؛ لا يُشترط استدعاء كل مكتبة مختارة في كل مهمة.'
if (@($result.connectionWarnings).Count -gt 0) {
    Write-Warning ($result.connectionWarnings -join '; ')
}
Write-Host ([Environment]::NewLine + 'الإجابة:')
Write-Output $result.answer
