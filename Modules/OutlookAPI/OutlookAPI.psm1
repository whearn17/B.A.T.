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

function Save-MailItemToPST {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$MailItem,

        [Parameter(Mandatory = $true)]
        [System.__ComObject]$TargetPST
    )

    # Access the "Threat Actor Accessed" folder within the provided PST
    $targetFolder = $TargetPST.GetRootFolder().Folders | Where-Object { $_.Name -eq "Threat Actor Accessed" }

    # Check if the folder exists
    if ($null -eq $targetFolder) {
        Write-Error "Error: Folder 'Threat Actor Accessed' not found in the provided PST."
        return
    }

    # Copy the mail item
    $copiedMailItem = $MailItem.Copy()

    # Move the mail item to the target folder ($null supresses console output)
    $null = $copiedMailItem.Move($targetFolder)

    # Release com object for copied mail item
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($copiedMailItem) | Out-Null
    $copiedMailItem = $null

    Start-Sleep -Milliseconds 100
}



function New-PST {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$OutlookApp,

        [Parameter(Mandatory = $true)]
        [string]$PSTDisplayName
    )

    # Check if PST already added in Outlook
    $PST = $OutlookApp.Session.Stores | Where-Object { $_.DisplayName -eq $PSTDisplayName }

    # Create PSTPath
    $PSTPath = "C:\temp\$($PSTDisplayName).pst"

    if ($null -eq $PST) {
        # Add the PST to Outlook
        $OutlookApp.Session.AddStoreEx($PSTPath, 2)  # 2 is for olStoreDefault, making it a default store type
    }

    # Brief pause to allow Outlook to fully add and initialize the PST
    Start-Sleep -Seconds 2

    # Re-query to get the newly added PST by its file path
    $PST = $OutlookApp.Session.Stores | Where-Object { $_.FilePath -eq $PSTPath }

    # Check if PST was successfully retrieved or created
    if ($null -eq $PST) {
        Write-Output "Error: Unable to find or create PST with display name: $PSTDisplayName"
        return
    }

    # Access the root folder of the PST
    $rootFolder = $PST.GetRootFolder()

    # Create a new folder in the PST file
    $null = $rootFolder.Folders.Add("Threat Actor Accessed")

    # Return the entire PST object
    return $PST
}