param (
    [Parameter(Mandatory = $false)]
    [string]$PSTDisplayName,

    [Parameter(Mandatory = $true)]
    [string]$MessageIDFile
)

Import-Module .\Modules\OutlookAPI\OutlookAPI.psm1
Import-Module .\Modules\SearchFunctions\SearchFunctions.psm1
Import-Module .\Modules\Logging\Logging.psm1

$applicationName = "BAT"
$environment = "Prod"
$startDateScript = Get-Date -Format "yyyy-MM-HH-mm-ss"

$global:LogFile = "$(Get-Location)\Logs\$($applicationName)_$($environment)_$($startDateScript).log"
function main {
    if (-not(Test-Path $global:LogFile)) {
        New-item -Path $global:LogFile | Out-Null
    }

    if (-not(Test-Path $MessageIDFile)) {
        throw "MessageIDFile does not exist -> $($MessageIDFile)"
    }

    $messageIDs = Get-Content $MessageIDFile

    if ([string]::IsNullOrEmpty($messageIDs)) {
        throw "MessageIDFile is empty"
    }

    Write-ScreenLog -Message "`n`nSearching for $(($messageIDs).Count) message ID(s)`n" -Level "info"
    Write-FileLog -Message "`n`nSearching for $(($messageIDs).Count) message ID(s)`n" -Level "info"
    
    $OutlookApp = New-OutlookComObject

    # If the Outlook GUI is open
    if ($OutlookApp.Explorers.Count -gt 0) {
        throw "Please close the Outlook GUI (No need to close background process)"
    }

    if ($null -eq $OutlookApp) {
        throw "Failed to create Outlook Com Object"
    }

    $PSTs = Get-OutlookConnectedPSTs $OutlookApp

    Write-ScreenLog -Message "`nAttached PSTs:`n$($PSTs | ForEach-Object { $_.DisplayName } | Out-String)" -Level "info"
    Write-FileLog -Message "`nAttached PSTs:`n$($PSTs | ForEach-Object { $_.DisplayName } | Out-String)" -Level "info"

    # If the user wants to save to a PST
    if (-not($PSTDisplayName -eq "")) {

        if (-not(Test-Path ".\Data")) {
            mkdir Data | Out-Null
        }

        Write-ScreenLog -Message "`nCreating new PST -> $($PSTDisplayName)" -Level info
        Write-FileLog -Message "`nCreating new PST -> $($PSTDisplayName)" -Level info

        $TargetPST = New-PST -OutlookApp $OutlookApp -PSTDisplayName $PSTDisplayName

        Search-ForMessageIDsInOutlook -PSTs $PSTs -TargetMessageIDs $messageIDs -TargetPST $TargetPST
    }
    # If the user does not want to save to a PST
    else {
        Search-ForMessageIDsInOutlook -PSTs $PSTs -TargetMessageIDs $messageIDs
    }   
}


try {
    main
}
catch {
    # Get the error type name and stack trace
    $errorTypeName = $_.Exception.GetType().FullName
    $errorStackTrace = $_.Exception.StackTrace

    # Create a detailed error message
    $detailedMessage = "Error Type: $errorTypeName`r`nMessage: $($_.Exception.Message)`r`nStack Trace: $errorStackTrace"

    # Log the detailed error message
    Write-ScreenLog -Message $detailedMessage -Level "fatal"
    Write-FileLog -Message $detailedMessage -Level "fatal"

    try {
        Remove-OutlookComObject $OutlookApp
    }
    catch {
        Write-ScreenLog -Message "Failed to remove Outlook COM Object: $($_.Exception.Message)" -Level "fatal"
        Write-FileLog -Message "Failed to remove Outlook COM Object: $($_.Exception.Message)" -Level "fatal"
        exit
    }
}
