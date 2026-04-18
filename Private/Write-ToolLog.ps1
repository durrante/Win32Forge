<#
.SYNOPSIS
    Structured logger for Win32Forge.

.DESCRIPTION
    Appends timestamped, caller-tagged log entries to the configured log file when
    VerboseLogging is enabled in $global:IntuneUploaderConfig. Silently no-ops when
    logging is disabled or unconfigured. Never throws — a log write failure never
    interrupts the main operation.

    Start-ToolLogSession writes a header separator to the log at tool startup.
#>

function Write-ToolLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',

        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )

    $cfg = $global:IntuneUploaderConfig
    if (-not $cfg -or -not $cfg.VerboseLogging) { return }
    $logPath = $cfg.LogPath
    if (-not $logPath) { return }

    $logDir = Split-Path $logPath -Parent -ErrorAction SilentlyContinue
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null
    }

    $ts     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $caller = (Get-PSCallStack)[1].Command
    if (-not $caller -or $caller -in '<ScriptBlock>', '') { $caller = 'Main' }

    $entry = "[$ts] [$Level] [$caller] $Message"

    if ($ErrorRecord) {
        $entry += "`r`n  Exception  : $($ErrorRecord.Exception.Message)"
        if ($ErrorRecord.Exception.InnerException) {
            $entry += "`r`n  Inner      : $($ErrorRecord.Exception.InnerException.Message)"
        }
        $entry += "`r`n  StackTrace : $($ErrorRecord.ScriptStackTrace)"
    }

    try { Add-Content -Path $logPath -Value $entry -Encoding UTF8 -ErrorAction Stop }
    catch {}
}

function Start-ToolLogSession {
    [CmdletBinding()]
    param([string]$Label = 'Session')

    $cfg = $global:IntuneUploaderConfig
    if (-not $cfg -or -not $cfg.VerboseLogging -or -not $cfg.LogPath) { return }

    $sep = '=' * 80
    $ts  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $hdr = @(
        ''
        $sep
        "[$ts] [$Label] Win32Forge — Log Session Start"
        "  PS Version : $($PSVersionTable.PSVersion)"
        "  Log Path   : $($cfg.LogPath)"
        $sep
    ) -join "`r`n"

    try { Add-Content -Path $cfg.LogPath -Value $hdr -Encoding UTF8 -ErrorAction Stop }
    catch {}
}
