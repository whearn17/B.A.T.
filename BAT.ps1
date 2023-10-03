param (
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory
)

. .\Search-Functions.ps1
. .\OutlookAPI.ps1

$messageIDs = Get-Content ".\mids.txt"

Write-Host "`n`nSearching for $(($messageIDs).Count) message ID(s)`n"


$OutlookApp = New-OutlookComObject

$PSTs = Get-OutlookConnectedPSTs $OutlookApp

foreach ($PST in $PSTs) {
    Search-ForMessageIDsInOutlook -PST $PST -TargetMessageIDs $messageIDs -OutputDirectory $OutputDirectory
}


Remove-OutlookComObject $OutlookApp