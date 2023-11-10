function Test-OutlookGUIOpen {
    try {
        $outlookProcesses = Get-Process | Where-Object { $_.ProcessName -like "OUTLOOK" }

        foreach ($process in $outlookProcesses) {
            if ($process.MainWindowTitle -ne "") {
                return $true
            }
        }

        return $false
    }
    catch {
        return $false
    }
}