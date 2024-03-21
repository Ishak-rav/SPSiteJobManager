# Function to install the module
function Install-ModuleIfNeeded {
    param (
        [string]$ModuleName
    )
    if (-not (ModuleChecker -ModuleName $ModuleName)) {
        Write-Host "The module '$ModuleName' is not installed. Attempting to install..."
        Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
        Write-Host "Module '$ModuleName' installed successfully."
    }
    else {
        Write-Host "Module '$ModuleName' is already installed."
    }
}
