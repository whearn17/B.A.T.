# PowerShell Script to extract Message-ID from an email header

# Sample header (you can replace this with your own)
$header = Get-Content .\header.txt -Raw
# Regex to match the Message-ID
$regex = 'Message-ID:\s*<([^>]+)>'

# Extract Message-ID using the regex
if ($header -match $regex) {
    $messageId = $matches[1]
    Write-Output "Found Message-ID: $messageId"
}
else {
    Write-Output "No Message-ID found."
}
