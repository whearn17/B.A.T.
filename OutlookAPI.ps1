function New-OutlookComObject {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    return $namespace
}

function Remove-OutlookComObject {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$ComObject
    )
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-OutlookConnectedPSTs {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$OutlookNamespace
    )

    $ConnectedPSTs = @()

    # Loop through all stores in the provided Outlook namespace
    foreach ($store in $OutlookNamespace.Stores) {
        $ConnectedPSTs += $store
    }

    return $ConnectedPSTs
}