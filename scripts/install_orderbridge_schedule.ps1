[CmdletBinding()]
param(
    [ValidateRange(0, 23)]
    [int]$StartHour = 7,
    [ValidateRange(0, 23)]
    [int]$EndHour = 21,
    [ValidateRange(0, 59)]
    [int]$StartMinute = 5,
    [ValidateRange(1, 24)]
    [int]$IntervalHours = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($EndHour -lt $StartHour) {
    throw 'EndHour must be at or after StartHour.'
}
$timeZone = Get-TimeZone
if ($timeZone.Id -ne 'SE Asia Standard Time') {
    throw "Windows must use SE Asia Standard Time (UTC+07:00) before installing this schedule; current zone is $($timeZone.Id)."
}

$projectRoot = Split-Path -Parent $PSScriptRoot
$powershellPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$serviceScript = Join-Path $PSScriptRoot 'start_orderbridge_service.ps1'
$workerScript = Join-Path $PSScriptRoot 'run_scheduled_workflow.ps1'
foreach ($scriptPath in @($serviceScript, $workerScript)) {
    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "Required scheduler script is missing: $scriptPath"
    }
}

$account = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $account -LogonType Interactive -RunLevel Limited
$serviceAction = New-ScheduledTaskAction -Execute $powershellPath `
    -Argument "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$serviceScript`""
$workerAction = New-ScheduledTaskAction -Execute $powershellPath `
    -Argument "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$workerScript`""

$serviceTrigger = New-ScheduledTaskTrigger -AtLogOn
$serviceSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
try {
    Register-ScheduledTask -TaskName 'OrderBridge DN2 - Service' -Description `
        'Starts the local OrderBridge service after the interactive user signs in.' `
        -Action $serviceAction -Trigger $serviceTrigger -Settings $serviceSettings -Principal $principal -Force | Out-Null

    $workerTriggers = @()
    for ($hour = $StartHour; $hour -le $EndHour; $hour += $IntervalHours) {
        $workerTriggers += New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours($hour).AddMinutes($StartMinute))
    }
    $workerSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -RunOnlyIfNetworkAvailable `
        -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 110)
    Register-ScheduledTask -TaskName 'OrderBridge DN2 - Every 2 Hours' -Description `
        'Exports the current Asia/Bangkok Sopoka report and imports only new, verified KiotViet orders.' `
        -Action $workerAction -Trigger $workerTriggers -Settings $workerSettings -Principal $principal -Force | Out-Null
} catch {
    throw "Windows could not register the Scheduled Tasks. Open Windows PowerShell as Administrator, then run this installer again from the Automation folder. Detail: $($_.Exception.Message)"
}

[PSCustomObject]@{
    service_task = 'OrderBridge DN2 - Service'
    workflow_task = 'OrderBridge DN2 - Every 2 Hours'
    schedule = ('{0:00}:{1:00} through {2:00}:{1:00}, every {3} hour(s), Asia/Bangkok' -f $StartHour, $StartMinute, $EndHour, $IntervalHours)
    user = $account
    run_only_when_logged_on = $true
} | ConvertTo-Json -Depth 3
