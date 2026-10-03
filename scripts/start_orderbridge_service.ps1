[CmdletBinding()]
param(
    [ValidateRange(1, 65535)]
    [int]$Port = 8765,
    [ValidateRange(5, 120)]
    [int]$StartupTimeoutSeconds = 45
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonPath = Join-Path $projectRoot '.venv\Scripts\python.exe'
$healthUrl = "http://127.0.0.1:$Port/api/state"

function Test-OrderBridgeHealth {
    try {
        $response = Invoke-RestMethod -Uri $healthUrl -TimeoutSec 3
        return $null -ne $response -and $response.ok -eq $true
    } catch {
        return $false
    }
}

if (Test-OrderBridgeHealth) {
    Write-Output "OrderBridge service is already ready on port $Port."
    exit 0
}
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Python environment was not found: $pythonPath"
}

Start-Process -FilePath $pythonPath `
    -ArgumentList @('app.py', '--host', '127.0.0.1', '--port', "$Port") `
    -WorkingDirectory $projectRoot `
    -WindowStyle Hidden

$deadline = (Get-Date).AddSeconds($StartupTimeoutSeconds)
do {
    Start-Sleep -Milliseconds 500
    if (Test-OrderBridgeHealth) {
        Write-Output "OrderBridge service started on port $Port."
        exit 0
    }
} while ((Get-Date) -lt $deadline)

throw "OrderBridge service did not become ready within $StartupTimeoutSeconds seconds."
