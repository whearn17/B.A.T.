$messageIDs = Get-Content .\failedMessageIDs.txt
$headerDump = Get-Content C:\temp\HeaderDump.txt -Raw

foreach ($messageID in $messageIDs) {
    if ($headerDump -match $messageID)
    {
        Write-Host "MessageID: $($messageID) found" -ForegroundColor Yellow
    }   
}