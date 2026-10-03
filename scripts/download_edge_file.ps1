[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DownloadDirectory,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ExpectedPrefix,

    [Parameter(Mandatory = $true)]
    [datetime]$NotBeforeUtc,

    [Parameter(Mandatory = $true)]
    [int]$BrowserProcessId,

    [ValidateRange(5, 300)]
    [int]$TimeoutSeconds = 150
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$directory = [IO.Path]::GetFullPath($DownloadDirectory)
if (-not [IO.Directory]::Exists($directory)) {
    throw "Download directory was not found: $directory"
}

function Find-EdgeDownloadButton {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        'Download file'
    )
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $BrowserProcessId
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $nameCondition,
        $processCondition
    )
    return $root.FindFirst([System.Windows.Automation.TreeScope]::Subtree, $condition)
}

function Get-CompletedReport {
    $candidate = Get-ChildItem -LiteralPath $directory -Filter "$ExpectedPrefix*" -File |
        Where-Object {
            $_.Length -gt 0 -and
            $_.LastWriteTimeUtc -ge $NotBeforeUtc -and
            $_.Extension.ToLowerInvariant() -in @('.xls', '.xlsx', '.xlsm', '.xltx', '.xltm')
        } |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1
    if ($null -eq $candidate) {
        return $null
    }
    if (Get-ChildItem -LiteralPath $directory -Filter '*.crdownload' -File -ErrorAction SilentlyContinue) {
        return $null
    }
    $length = $candidate.Length
    Start-Sleep -Milliseconds 750
    $fresh = Get-Item -LiteralPath $candidate.FullName -ErrorAction SilentlyContinue
    if ($null -eq $fresh -or $fresh.Length -ne $length) {
        return $null
    }
    return $fresh
}

$deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$clicked = $false
while ([DateTime]::UtcNow -lt $deadline) {
    $ready = Get-CompletedReport
    if ($null -ne $ready) {
        [ordered]@{
            success = $true
            path    = $ready.FullName
            size    = [long]$ready.Length
            clicked = $clicked
        } | ConvertTo-Json -Compress
        exit 0
    }

    if (-not $clicked) {
        $button = Find-EdgeDownloadButton
        if ($null -ne $button -and $button.Current.IsEnabled) {
            $pattern = $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            ([System.Windows.Automation.InvokePattern]$pattern).Invoke()
            $clicked = $true
        }
    }
    Start-Sleep -Milliseconds 250
}

if (-not $clicked) {
    throw 'Microsoft Edge did not show its Download file button for the exported Sopoka report.'
}
throw "Edge accepted Download file but did not finish an Excel report ($ExpectedPrefix*.xls/.xlsx/.xlsm) within $TimeoutSeconds seconds."
