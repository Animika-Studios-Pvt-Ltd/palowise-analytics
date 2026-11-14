# ==============================================================================
# Palowise Excel-to-PostgreSQL Recurring Sync Runner
# Executed by Windows Task Scheduler or manual trigger.
# ==============================================================================

$ErrorActionPreference = "Continue"

$ScriptDir   = $PSScriptRoot
$VenvPython  = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$SyncScript  = Join-Path $ScriptDir "sync_excel_to_pg.py"
$LogFile     = Join-Path $ScriptDir "sync.log"

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Helper function to append lines using strict UTF-8 encoding
function Write-SyncLog {
    param([string]$Message)
    [System.IO.File]::AppendAllText($LogFile, "$Message`r`n", [System.Text.Encoding]::UTF8)
}

Write-SyncLog "[$Timestamp] Starting Palowise recurring sync job..."

if (-not (Test-Path $VenvPython)) {
    Write-SyncLog "[$Timestamp] ERROR: Virtual environment Python not found at $VenvPython"
    exit 1
}

# Execute recurring sync, capturing stdout and stderr cleanly without PowerShell encoding corruption
$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
$ProcessInfo.FileName               = $VenvPython
$ProcessInfo.Arguments              = "`"$SyncScript`" --recurring"
$ProcessInfo.WorkingDirectory       = $ScriptDir
$ProcessInfo.RedirectStandardOutput = $true
$ProcessInfo.RedirectStandardError  = $true
$ProcessInfo.UseShellExecute        = $false
$ProcessInfo.CreateNoWindow         = $true
$ProcessInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
$ProcessInfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

$Process = New-Object System.Diagnostics.Process
$Process.StartInfo = $ProcessInfo

$stdout = ""
$stderr = ""

try {
    $null = $Process.Start()
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    $Process.WaitForExit()
    $ExitCode = $Process.ExitCode
}
catch {
    $stderr = $_.Exception.Message
    $ExitCode = 1
}

if ($stdout.Trim().Length -gt 0) {
    Write-SyncLog $stdout.TrimEnd()
}

if ($stderr.Trim().Length -gt 0) {
    Write-SyncLog "STDERR: $($stderr.TrimEnd())"
}

$EndTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-SyncLog "[$EndTimestamp] Finished with exit code: $ExitCode"
Write-SyncLog ("-" * 70)

exit $ExitCode
