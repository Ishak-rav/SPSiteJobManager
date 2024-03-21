# Function to import the module
function Import-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    if (-not (ModuleChecker -ModuleName $ModuleName)) {
        Write-Host "Importing module '$ModuleName'..."
        Import-Module $ModuleName
        Write-Host "Module '$ModuleName' imported successfully."
    }
    else {
        Write-Host "Module '$ModuleName' is already imported."
    }
}
