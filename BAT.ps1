param (
    [Parameter(Mandatory = $false)]
    [string]$PSTDisplayName,

    [Parameter(Mandatory = $false)]
    [string]$MessageIDFile,

    [Parameter(Mandatory = $false)]
    [string]$SubjectNameFile,

    [Parameter(Mandatory = $false)]
    [switch]$Help
)

Import-Module .\Modules\OutlookAPI\OutlookAPI.psm1
Import-Module .\Modules\SearchFunctions\SearchFunctions.psm1
Import-Module .\Modules\Logging\Logging.psm1
Import-Module .\Modules\Utils\SessionMeta.psm1
Import-Module .\Modules\Utils\HWNDEx.psm1

$applicationName = "BAT"
$environment = "Prod"
$startDateScript = Get-Date -Format "yyyy_MM_dd-HH-mm-ss"

$global:LogFile = "$(Get-Location)\Logs\$($applicationName)_$($environment)_$($startDateScript).log"

function MessageIDSearch {
    $messageIDs = Get-Content $MessageIDFile

    if ([string]::IsNullOrEmpty($messageIDs)) {
        throw "MessageIDFile is empty"
    }

    Write-ScreenLog -Message "`n`nSearching for $(($messageIDs).Count) message ID(s)`n" -Level "info"
    Write-FileLog -Message "`n`nSearching for $(($messageIDs).Count) message ID(s)`n" -Level "info"

    $OutlookApp = New-OutlookComObject

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

function SubjectNameSearch {
    $subjectNames = Get-Content $SubjectNameFile

    if ([string]::IsNullOrEmpty($subjectNames)) {
        throw "SubjectNameFile is empty"
    }

    Write-ScreenLog -Message "`n`nSearching for $(($subjectNames).Count) subject names(s)`n" -Level "info"
    Write-FileLog -Message "`n`nSearching for $(($subjectNames).Count) subject names(s)`n" -Level "info"

    $OutlookApp = New-OutlookComObject

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

        Get-MessagesBySubject -PSTs $PSTs -TargetSubjects $subjectNames -TargetPST $TargetPST
    }
    # If the user does not want to save to a PST
    else {
        Get-MessagesBySubject -PSTs $PSTs -TargetSubjects $subjectNames
    }   
}

function Usage {
    Write-Host "Usage:"
    Write-Host "--------------------------------------"
    Write-Host "This script is designed to search for emails in Outlook based on Message IDs or Subject Names."
    Write-Host "It can be used to save search results into a new PST file if desired."
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -PSTDisplayName: The display name for a new PST file where results will be saved. This parameter is optional."
    Write-Host "  -MessageIDFile: The path to a file containing Message IDs to search for. This parameter is optional."
    Write-Host "  -SubjectNameFile: The path to a file containing Subject Names to search for. This parameter is optional."
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\BAT.ps1 -PSTDisplayName 'MyPSTFile' -MessageIDFile 'C:\path\to\messageids.txt'"
    Write-Host "  .\BAT.ps1 -SubjectNameFile 'C:\path\to\subjectnames.txt'"
    Write-Host ""
    Write-Host "Note: Do not run this script with both MessageIDFile and SubjectNameFile at the same time."
    Write-Host "Note: Ensure Outlook is not running in GUI mode when executing this script."
}


function Main {
    if (Test-IsAdmin) {
        throw "Please don't run this program in an Administrator shell"
    }

    if (Test-OutlookGUIOpen) {
        throw "Please close the Outlook GUI before running this program. You do not need to kill through task manager."
    }

    if (([string]::IsNullOrEmpty($MessageIDFile) -and [string]::IsNullOrEmpty($SubjectNameFile) -and [string]::IsNullOrEmpty($PSTDisplayName)) -or $Help) {
        Usage
    }

    if (-not(Test-Path $global:LogFile)) {
        New-item -Path $global:LogFile | Out-Null
    }

    if ((-not [string]::IsNullOrEmpty($MessageIDFile)) -and (-not(Test-Path $MessageIDFile))) {
        throw "Supplied MessageIDFile does not exist"
    }

    if (-not [string]::IsNullOrEmpty($MessageIDFile) -and -not [string]::IsNullOrEmpty($SubjectNameFile)) {
        throw "Please don't supply a MessageIDFile and a SubjectNameFile"
    }    
    
    if (-not [string]::IsNullOrEmpty($MessageIDFile)) {
        MessageIDSearch
    }

    if (-not [string]::IsNullOrEmpty($SubjectNameFile)) {
        SubjectNameSearch
    }
}


try {
    Main
}
catch {
    # Create a detailed error message
    Write-Host "$($_.InvocationInfo.ScriptLineNumber))"
    $detailedMessage = "$($_.Exception.Message)"

    # Log the detailed error message
    Write-ScreenLog -Message $detailedMessage -Level "fatal"
    Write-FileLog -Message $detailedMessage -Level "fatal"

    try {
        Remove-OutlookComObject $OutlookApp
    }
    catch {
        exit
    }
}
