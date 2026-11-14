# ==============================================================================
# Palowise Supabase Task Scheduler Registration Script
# Registers a Windows Scheduled Task to run every 30 minutes.
# ==============================================================================

param (
    [switch]$Register = $false,
    [switch]$Unregister = $false
)

$TaskName     = "Palowise_Supabase_Sync"
$ScriptDir    = $PSScriptRoot
$RunnerScript = Join-Path $ScriptDir "run_supabase_sync.ps1"

if ($Unregister) {
    Write-Host "Unregistering scheduled task '$TaskName'..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Task unregistered successfully."
    exit 0
}

if (-not $Register) {
    Write-Host "======================================================================"
    Write-Host "Palowise Supabase Task Scheduler Registration Template"
    Write-Host "======================================================================"
    Write-Host "Task Name:        $TaskName"
    Write-Host "Runner Script:    $RunnerScript"
    Write-Host "Interval:         Every 30 minutes"
    Write-Host "Execution Policy: Bypass"
    Write-Host "Instance Policy:  IgnoreNew (Disallow overlapping executions)"
    Write-Host ""
    Write-Host "To register this task, run:"
    Write-Host "  powershell -ExecutionPolicy Bypass -File .\setup_supabase_scheduler.ps1 -Register"
    Write-Host "======================================================================"
    exit 0
}

# 1. Action: Execute PowerShell runner script
$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$RunnerScript`"" `
    -WorkingDirectory $ScriptDir

# 2. Trigger: Daily, repeat every 30 minutes indefinitely
$Time = (Get-Date).AddMinutes(5).ToString("HH:mm")
$Trigger = New-ScheduledTaskTrigger -Daily -At $Time
$Trigger.Repetition = (New-ScheduledTaskTrigger -Once -At $Time -RepetitionInterval (New-TimeSpan -Minutes 30)).Repetition

# 3. Settings: Disallow overlapping instances, stop if runs longer than 10 mins
$Settings = New-ScheduledTaskSettingsSet `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries

# 4. Principal: Current user interactive session
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive

# 5. Register Task
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Automated 30-minute Palowise marketing analysis sync from OneDrive Excel to Supabase Cloud Database" `
    -Force

Write-Host "SUCCESS: Scheduled task '$TaskName' registered successfully."
Write-Host "Next run scheduled for: $Time (repeats every 30 minutes)."

