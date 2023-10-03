function Get-PSTFoldersRecursive {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$Folder
    )

    $folders = @($Folder)  # Initialize with the current folder

    foreach ($subfolder in $Folder.Folders) {
        # Recursively get subfolders and add to the array
        $folders += Get-PSTFoldersRecursive -Folder $subfolder
    }

    return $folders
}

function Save-MailItemAsMsg {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$MailItem,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    # Create a unique filename for the mail item
    $filenameBase = ("{0}-{1}" -f $MailItem.Subject, $MailItem.EntryID) -replace '[\W]', '_'
    $filename = "$filenameBase.msg"

    $filepath = Join-Path -Path $OutputPath -ChildPath $filename
    $MailItem.SaveAs($filepath, 3)  # Use 3, which is the enumeration value for olMSG
}

function Search-ForMessageIDsInOutlook {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$PST,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetMessageIDs,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )

    $rootFolder = $PST.GetRootFolder()

    # Get all folders, including nested ones
    $allFolders = Get-PSTFoldersRecursive -Folder $rootFolder

    $MessageIDPattern = "Message-ID:\s*<([^>]+)>"

    # Create a hashtable for faster lookup
    $targetIDsHash = @{}
    foreach ($id in $TargetMessageIDs) {
        $targetIDsHash[$id.ToLower()] = $false  # Initialize with 'false' indicating not yet found
    }

    # Ensure the output directory exists
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory
    }

    # Iterate over all folders
    foreach ($folder in $allFolders) {
        # Iterate over each mail item in the folder
        foreach ($mail in $folder.Items) {
            # Check if the item is an email or calendar invite
            if ($mail.MessageClass -eq "IPM.Note" -or $mail.MessageClass -like "IPM.Schedule.Meeting.*") {
                # Extract the header and search for Message-ID
                $header = $mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                if ($header -match $MessageIDPattern) {
                    $foundID = $matches[1].ToLower()  # Convert found ID to lowercase
                    # Check if found ID matches any in the provided list
                    if ($targetIDsHash.ContainsKey($foundID)) {
                        Write-Host "[+] Matched Message-ID: $foundID in folder: $($folder.Name)" -ForegroundColor Green

                        Save-MailItemAsMsg -MailItem $mail -OutputPath $OutputDirectory
                        
                        $targetIDsHash[$foundID] = $true  # Mark as found
                    }
                }
            }
        }
    }
    
    # Report any MessageIDs that weren't found
    $notFoundMessageIDs = $targetIDsHash.Keys | Where-Object { $targetIDsHash[$_] -eq $false }
    if ($notFoundMessageIDs.Count -gt 0) {
        Write-Host "The following Message-IDs were not found:" -ForegroundColor Yellow
        $notFoundMessageIDs | ForEach-Object {
            Write-Host "[-] $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "All provided Message-IDs were found." -ForegroundColor Green
    }
}
