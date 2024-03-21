# Function to check if the module is available
function ModuleChecker {
    param (
        [string]$ModuleName
    )
    return Get-Module -ListAvailable -Name $ModuleName
}
