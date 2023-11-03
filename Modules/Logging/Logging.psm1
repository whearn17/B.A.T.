function Write-ScreenLog {
    param (
        [String]$Message,
        [String]$Level
    )

    switch ($Level) {
        "debug" { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [DEBUG]    $Message" -ForegroundColor Gray }
        "info" { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO]     $Message" -ForegroundColor White }
        "warning" { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WARNING]  $Message" -ForegroundColor Yellow }
        "error" { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR]    $Message" -ForegroundColor Red }
        "fatal" { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [FATAL]    $Message" -ForegroundColor DarkRed }
        default { Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [UNKNOWN]  $Message" -ForegroundColor Magenta }
    }
}

function Write-FileLog {
    param (
        [String]$Message,
        [String]$Level
    )

    # Construct the log message similar to Write-ScreenLog
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    switch ($Level) {
        "debug" { $logMessage = "$timestamp [DEBUG]    $Message" }
        "info" { $logMessage = "$timestamp [INFO]     $Message" }
        "warning" { $logMessage = "$timestamp [WARNING]  $Message" }
        "error" { $logMessage = "$timestamp [ERROR]    $Message" }
        "fatal" { $logMessage = "$timestamp [FATAL]    $Message" }
        default { $logMessage = "$timestamp [UNKNOWN]  $Message" }
    }

    # Append the log message to the file
    Add-Content -Path $global:LogFile -Value $logMessage
}
