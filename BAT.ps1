param (
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $true)]
    [string]$MessageIDFile
)

. .\Search-Functions.ps1
. .\OutlookAPI.ps1

$messageIDs = Get-Content $MessageIDFile

Write-Host "`n`nSearching for $(($messageIDs).Count) message ID(s)`n"


$OutlookApp = New-OutlookComObject

$PSTs = Get-OutlookConnectedPSTs $OutlookApp

Write-Host "`nAttached PSTs:`n$($PSTs | ForEach-Object { $_.DisplayName } | Out-String)"

Search-ForMessageIDsInOutlook -PST $PSTs -TargetMessageIDs $messageIDs -OutputDirectory $OutputDirectory


Remove-OutlookComObject $OutlookApp