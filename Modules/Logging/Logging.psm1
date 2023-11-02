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

}