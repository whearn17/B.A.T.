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

function Search-ForMessageIDsInOutlook {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject[]]$PSTs,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetMessageIDs,

        [Parameter(Mandatory = $false)]
        [System.__ComObject]$TargetPST
    )

    $MessageIDPattern = "(?i)Message-ID:\s*<([^>]+)>"

    # Create a hashtable for faster lookup
    $targetIDsHash = @{}
    foreach ($id in $TargetMessageIDs) {
        $targetIDsHash[$id.ToLower()] = $false  # Initialize with 'false' indicating not yet found
    }

    # Iterate over all provided PSTs
    foreach ($PST in $PSTs) {
        Write-FileLog -Message "Working on PST -> $($PST.DisplayName)" -Level "info"

        $rootFolder = $PST.GetRootFolder()
        $allFolders = Get-PSTFoldersRecursive -Folder $rootFolder

        # Iterate over all folders
        foreach ($folder in $allFolders) {
            Write-FileLog -Message "Working on folder $($folder.DisplayName)" -Level "info"
            # Iterate over each mail item in the folder
            foreach ($mail in $folder.Items) {

                # Check to make sure the mail item has a message header
                if ($null -eq $mail.PropertyAccessor) { continue }

                # Extract the header and search for Message-ID
                $header = $mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                
                $matchedIDs = [regex]::Matches($header, $MessageIDPattern)
                
                foreach ($match in $matchedIDs) {
                    $foundID = $match.Groups[1].Value.ToLower()

                    # Check if found ID matches any in the provided list
                    if ($targetIDsHash.ContainsKey($foundID)) {
                        Write-ScreenLog -Message "[+] Matched Message-ID: $foundID with Subject: $($mail.Subject)" -Level "info"
                        Write-FileLog -Message "[+] Matched Message-ID: $foundID with Subject: $($mail.Subject)" -Level "info"

                        $targetIDsHash[$foundID] = $true  # Mark as found

                        # Check to make sure the user passed in a targetPST to save to
                        if (-not($null -eq $TargetPST)) {
                            Save-MailItemToPST -MailItem $mail -TargetPST $TargetPST
                        }
                    }
                }

                # Release com object for mail item
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
                $mail = $null
            }            
        }
    }
    
    # Report any MessageIDs that weren't found
    $notFoundMessageIDs = $targetIDsHash.Keys | Where-Object { $targetIDsHash[$_] -eq $false }
    if ($notFoundMessageIDs.Count -gt 0) {
        Write-ScreenLog -Message "The following Message-IDs were not found:" -Level "warning"
        Write-FileLog -Message "The following Message-IDs were not found:" -Level "warning"
        $notFoundMessageIDs | ForEach-Object {
            Write-ScreenLog -Message "[-] $_" -Level "warning"
            Write-FileLog -Message "[-] $_" -Level "warning"
        }
    }
    else {
        Write-ScreenLog -Message "All provided Message-IDs were found." -Level "info"
        Write-FileLog -Message "All provided Message-IDs were found." -Level "info"
    }
}

function Get-NormalizedSubject {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Subject
    )

    # Removing special characters using regex and converting to lowercase
    return ($Subject -replace '[\W]', '').ToLower()
}

function Get-MessagesBySubject {
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject[]]$PSTs,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetSubjects,

        [Parameter(Mandatory = $false)]
        [System.__ComObject]$TargetPST
    )

    # Normalize target subjects for comparison
    $targetSubjectsHash = @{}
    foreach ($subject in $TargetSubjects) {
        $normalizedSubject = Get-NormalizedSubject -Subject $subject
        $targetSubjectsHash[$normalizedSubject] = $false  # Initialize with 'false' indicating not yet found
    }
    
    # Iterate over all provided PSTs
    foreach ($PST in $PSTs) {
        $rootFolder = $PST.GetRootFolder()

        # Get all folders, including nested ones
        $allFolders = Get-PSTFoldersRecursive -Folder $rootFolder

        # Iterate over all folders
        foreach ($folder in $allFolders) {
            foreach ($mail in $folder.Items) {
                if ("" -eq $mail.Subject) { continue }

                $currentSubject = Get-NormalizedSubject -Subject $mail.Subject
        
                if (-not $targetSubjectsHash.ContainsKey($currentSubject)) { continue }
        
                Write-ScreenLog -Message "[+] Matched Subject: $currentSubject" -Level "info"
                Write-FileLog -Message "[+] Matched Subject: $currentSubject" -Level "info"

                $targetSubjectsHash[$currentSubject] = $true

                # Check to make sure the user passed in a targetPST to save to
                if (-not($null -eq $TargetPST)) {
                    Save-MailItemToPST -MailItem $mail -TargetPST $TargetPST
                }                
            }            
        }        
    }

        
    # Report any subjects that weren't found
    $notFoundSubjects = $targetSubjectsHash.Keys | Where-Object { $targetSubjectsHash[$_] -eq $false }
    if ($notFoundSubjects.Count -gt 0) {
        Write-ScreenLog -Message "The following subjects were not found:" -Level "warning"
        Write-FileLog -Message "The following subjects were not found:" -Level "warning"
        $notFoundSubjects | ForEach-Object {
            Write-ScreenLog -Message "[-] $_" -Level Warning
            Write-FileLog -Message "[-] $_" -Level Warning
        }
    }
    else {
        Write-ScreenLog -Message "All provided subjects were found and their corresponding Message-IDs written to $outputFile." -Level "info"
        Write-FileLog -Message "All provided subjects were found and their corresponding Message-IDs written to $outputFile." -Level "info"
    }
}

function Get-MessageByItemID {
    # Unifinished
    param (
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$OutlookApp,

        [Parameter(Mandatory = $true)]
        [string[]]$TargetItemIDs,

        [Parameter(Mandatory = $true)]
        [System.__ComObject]$TargetPST
    )

    foreach ($TargetItemID in $TargetItemIDs) {
        $mailItem = $OutlookApp.GetItemFromID($TargetItemID)
        
        Save-MailItemToPST -MailItem $mailItem -TargetPST $TargetPST

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mailItem) | Out-Null
        $mailItem = $null
    }

}