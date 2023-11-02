param (
    [Parameter(Mandatory = $true)]
    [string]$PSTPath
)

# Importing necessary libraries
Add-type -assembly "Microsoft.Office.Interop.Outlook"

Import-Module .\OutlookAPI.psm1

# Checking if directory exists
if (-not (Test-Path $PSTPath)) {
    Write-Host "Directory not found!"
    exit
}

$OutlookApp = New-OutlookComObject

# Getting all the .pst files in the directory
$pstFiles = Get-ChildItem -Path $PSTPath -Filter "*.pst"

# Checking if there are no .pst files in the directory
if ($pstFiles.Count -eq 0) {
    Write-Host "No PST files found in the given directory!"
    exit
}

# Looping through each PST file and adding it to Outlook
foreach ($pst in $pstFiles) {
    try {
        $OutlookApp.AddStore($pst.FullName)
        Write-Host "$($pst.Name) imported successfully!"
    }
    catch {
        Write-Host "Failed to import $($pst.Name). Error: $_"
    }
}

Remove-OutlookComObject $OutlookApp
