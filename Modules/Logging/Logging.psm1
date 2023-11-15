function Write-ScreenLog {
    param (
        [String]$Message,
        [String]$Level
    )

    # Store the formatted date in a variable
    $currentDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    switch ($Level) {
        "debug" { Write-Host "$($currentDate) [DEBUG]    $Message" -ForegroundColor Gray }
        "info" { Write-Host "$($currentDate) [INFO]     $Message" -ForegroundColor White }
        "warning" { Write-Host "$($currentDate) [WARNING]  $Message" -ForegroundColor Yellow }
        "error" { Write-Host "$($currentDate) [ERROR]    $Message" -ForegroundColor Red }
        "fatal" { Write-Host "$($currentDate) [FATAL]    $Message" -ForegroundColor DarkRed }
        default { Write-Host "$($currentDate) [UNKNOWN]  $Message" -ForegroundColor Magenta }
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
