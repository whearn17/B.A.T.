param (
    [Parameter(Mandatory = $false)]
    [string]$PSTDisplayName,

    [Parameter(Mandatory = $true)]
    [string]$MessageIDFile
)

Import-Module .\Modules\OutlookAPI\OutlookAPI.psm1
Import-Module .\Modules\SearchFunctions\SearchFunctions.psm1
Import-Module .\Modules\Logging\Logging.psm1
function main {
    if (-not(Test-Path $MessageIDFile)) {
        throw "MessageIDFile does not exist -> $($MessageIDFile)"
    }

    $messageIDs = Get-Content $MessageIDFile

    if ([string]::IsNullOrEmpty($messageIDs)) {
        throw "MessageIDFile is empty"
    }

    Write-ScreenLog -Message "`n`nSearching for $(($messageIDs).Count) message ID(s)`n" -Level "info"
    
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

    # If the user wants to save to a PST
    if (-not($PSTDisplayName -eq "")) {

        if (-not(Test-Path ".\Data")) {
            mkdir Data | Out-Null
        }

        Write-ScreenLog -Message "`nCreating new PST -> $($PSTDisplayName)" -Level info

        $TargetPST = New-PST -OutlookApp $OutlookApp -PSTDisplayName $PSTDisplayName

        Search-ForMessageIDsInOutlook -PSTs $PSTs -TargetMessageIDs $messageIDs -TargetPST $TargetPST
    }
    # If the user does not want to save to a PST
    else {
        Search-ForMessageIDsInOutlook -PSTs $PSTs -TargetMessageIDs $messageIDs
    }   
}


# try {
#     main 
# }
# catch {

#     Write-ScreenLog -Message "$($_)" -Level "fatal"
#     try {
#         Remove-OutlookComObject $OutlookApp
#     }
#     catch {
#         exit
#     }
    
# }
main