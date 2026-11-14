# ==============================================================================
# Palowise Excel-to-Supabase Cloud Synchronization Runner
# Executed by Windows Task Scheduler or manual invocation.
# ==============================================================================

param (
    [switch]$DryRun = $false,
    [switch]$Force = $false
)

$ErrorActionPreference = "Continue"

$ScriptDir    = $PSScriptRoot
$VenvPython   = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$SyncScript   = Join-Path $ScriptDir "sync_excel_to_supabase.py"
$LogFile      = Join-Path $ScriptDir "sync_supabase.log"

$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Helper function to append lines using strict UTF-8 encoding
function Write-SyncLog {
    param([string]$Message)
    [System.IO.File]::AppendAllText($LogFile, "$Message`r`n", [System.Text.Encoding]::UTF8)
}

# Determine script arguments
$ScriptArgs = "`"$SyncScript`""
if ($DryRun) {
    $ScriptArgs += " --dry-run"
    $ModeLabel = "DRY-RUN (Safe validation only)"
} elseif ($Force) {
    $ScriptArgs += " --recurring --force"
    $ModeLabel = "RECURRING SYNC (Forced)"
} else {
    $ScriptArgs += " --recurring"
    $ModeLabel = "RECURRING SYNC"
}

Write-SyncLog "[$Timestamp] Starting Palowise Supabase sync job ($ModeLabel)..."

if (-not (Test-Path $VenvPython)) {
    Write-SyncLog "[$Timestamp] ERROR: Virtual environment Python not found at $VenvPython"
    exit 1
}

if (-not (Test-Path $SyncScript)) {
    Write-SyncLog "[$Timestamp] ERROR: Sync script not found at $SyncScript"
    exit 1
}

# Execute process capturing stdout and stderr cleanly with UTF-8
$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
$ProcessInfo.FileName               = $VenvPython
$ProcessInfo.Arguments              = $ScriptArgs
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
if ($ExitCode -eq 0) {
    Write-SyncLog "[$EndTimestamp] Finished with exit code: 0 (SUCCESS)"
} else {
    Write-SyncLog "[$EndTimestamp] ERROR: Finished with exit code: $ExitCode (FAILURE)"
}
Write-SyncLog ("-" * 70)

exit $ExitCode

