[CmdletBinding()]
param(
    [string]$RepoPath = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$TaskName = "Heijunka Dashboard Update",
    [string[]]$RunTimes = @("08:00", "12:30", "16:30"),
    [switch]$ConfigurePower
)
$ErrorActionPreference = "Stop"
$repo = Resolve-Path $RepoPath
$runAll = Join-Path $repo "run_all.py"
if (-not (Test-Path $runAll)) {
    throw "Could not find run_all.py in RepoPath: $repo"
}
$batPath = Join-Path $repo "run_dashboard_update.bat"
$logPath = Join-Path $repo "run_all.log"
$batContent = @"
@echo off
cd /d "$repo"
python run_all.py >> "$logPath" 2>&1
"@
Set-Content -Path $batPath -Value $batContent -Encoding ASCII
$triggers = foreach ($time in $RunTimes) {
    try {
        $parsed = [datetime]::ParseExact($time, "HH:mm", [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "RunTimes values must use 24-hour HH:mm format, for example 08:00 or 16:30. Invalid value: $time"
    }
    New-ScheduledTaskTrigger -Daily -At $parsed.TimeOfDay
}
$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batPath`"" -WorkingDirectory $repo
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -WakeToRun `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
$task = New-ScheduledTask -Action $action -Trigger $triggers -Settings $settings -Principal $principal
Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
if ($ConfigurePower) {
    Write-Host "Configuring plugged-in power settings so the computer does not sleep..."
    powercfg /change standby-timeout-ac 0 | Out-Null
    powercfg /change hibernate-timeout-ac 0 | Out-Null
    powercfg /change monitor-timeout-ac 30 | Out-Null
}
Write-Host "Created/updated: $batPath"
Write-Host "Created/updated scheduled task: $TaskName"
Write-Host "Run times: $($RunTimes -join ', ')"
Write-Host "Log file: $logPath"
Write-Host "To test now, run: Start-ScheduledTask -TaskName '$TaskName'"