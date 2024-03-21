# Function to get the credentials with retries
function Get-CredentialWithRetries {
    [CmdletBinding()]
    param (
        [int]$MaxAttempts = 3
    )
    $attempts = 0
    while ($attempts -lt $MaxAttempts) {
        Write-Host "Please enter your credentials to continue. Attempt $([int]$attempts + 1) of $MaxAttempts."
        $cred = Get-Credential

        if ($cred) {
            return $cred
        }
        else {
            Write-Host "No credentials provided."
        }

        $attempts++
    }

    Write-Host "No identifier provided after $MaxAttempts attempts. The script will end."
    exit
}
