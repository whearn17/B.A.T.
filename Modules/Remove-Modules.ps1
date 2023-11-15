function Remove-Modules {
    # Get all .psm1 files in the current directory and its subdirectories
    $moduleFiles = Get-ChildItem -Path . -Recurse -Filter *.psm1

    foreach ($file in $moduleFiles) {
        # Get the module name from the file name
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file.FullName)

        # Check if the module is already loaded
        if (Get-Module -Name $moduleName) {
            # Remove the module
            Remove-Module -Name $moduleName
        }
    }
}

Remove-Modules