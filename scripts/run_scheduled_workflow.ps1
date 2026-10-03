[CmdletBinding()]
param(
    [ValidateRange(1, 65535)]
    [int]$Port = 8765
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$serviceStarter = Join-Path $PSScriptRoot 'start_orderbridge_service.ps1'
$logDirectory = Join-Path $projectRoot 'logs\scheduler'
$logPath = Join-Path $logDirectory ((Get-Date).ToString('yyyy-MM-dd') + '.jsonl')
$tickUrl = "http://127.0.0.1:$Port/api/scheduler/tick"

New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null

try {
    & $serviceStarter -Port $Port
    $response = Invoke-RestMethod -Uri $tickUrl -Method Post -ContentType 'application/json' `
        -Body '{}' -TimeoutSec 660
    $record = [ordered]@{
        recorded_at = (Get-Date).ToUniversalTime().ToString('o')
        task = 'OrderBridge DN2 - Every 2 Hours'
        response = $response
    }
    ($record | ConvertTo-Json -Depth 14 -Compress) | Add-Content -LiteralPath $logPath -Encoding utf8
    if ($response.ok -ne $true) {
        $message = [string]$response.message
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = 'Scheduled workflow stopped safely; review the scheduler log and dashboard.'
        }
        Write-Warning $message
        exit 2
    }
    $outcome = [string]$response.outcome
    if ([string]::IsNullOrWhiteSpace($outcome)) { $outcome = 'unknown' }
    Write-Output ("Scheduled workflow outcome: " + $outcome)
    exit 0
} catch {
    $record = [ordered]@{
        recorded_at = (Get-Date).ToUniversalTime().ToString('o')
        task = 'OrderBridge DN2 - Every 2 Hours'
        error = $_.Exception.Message
    }
    ($record | ConvertTo-Json -Depth 6 -Compress) | Add-Content -LiteralPath $logPath -Encoding utf8
    Write-Error $_
    exit 1
}
