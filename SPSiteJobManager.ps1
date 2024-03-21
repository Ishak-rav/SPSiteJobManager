# Base variables definition
$moduleName = 'PnP.PowerShell'
$startTime = Get-Date
$startTimeStr = $startTime.ToString("dd-MM-yyyy HH:mm:ss")

$path = "C:\Users\ichennouf\OneDrive - ELIADIS\Desktop\"
$exportPath = "${path}scripts\Resultat\"
$csvPath = "${exportPath}listSites.csv"
$basePath = "${exportPath}ExportCSV\"

# Importing functions
. "${path}SPSiteJobManager\ModuleChecker.ps1"
. "${path}SPSiteJobManager\InstallModuleIfNeeded.ps1"
. "${path}SPSiteJobManager\ImportModuleIfNeeded.ps1"
. "${path}SPSiteJobManager\ProcessSite.ps1"
. "${path}SPSiteJobManager\GetCredentialWithRetries.ps1"

# Initializing lists to store information
$allSites = @()
$allLibraries = @()
$allFiles = @()

Write-Host "-------------------------- Script start at $startTimeStr --------------------------"

# Ensure module is installed and imported
Install-ModuleIfNeeded -ModuleName $moduleName
Import-ModuleIfNeeded -ModuleName $moduleName

# Importing CSV
try {
    $listeSites = Import-Csv -Path $csvPath
    Write-Host "CSV import successful."
}
catch [System.IO.IOException] {
    Write-Host "CSV file access error: $_"
    exit
}
catch [System.TimeoutException] {
    Write-Host "Operation timed out: $_"
    exit
}
catch {
    Write-Host "General error when importing CSV: $_"
    exit
}

# Asking for credentials once
$cred = Get-CredentialWithRetries -MaxAttempts 3
if (-not $cred) {
    Write-Host "Script terminating due to lack of valid credentials."
    exit
}

# Starting a job for each site
$listeSites | ForEach-Object -Parallel {
    . "${using:path}SPSiteJobManager\ProcessSite.ps1"
    ProcessSite -site $_ -cred $using:cred
} -ThrottleLimit 10  

Write-Host "Waiting for all jobs to finish..."
# Waiting for all jobs to complete
Get-Job | Wait-Job
Write-Host "All jobs are completed"

# Retrieving results from each job
$jobs = Get-Job

# Job processing
foreach ($job in $jobs) {

    Write-Host "Processing job results $($job.Id)"
    $result = Receive-Job -Job $job

    if ($job.State -eq "Failed") {
        Write-Host "Job $($job.Id) failed"
        continue
    }

    # Checking that the result is not null
    if ($result) {
        # Adding the results to the global lists if they are present, independently of each other
        if ($result.Sites) {
            $allSites += $result.Sites
        }
        if ($result.Libraries) {
            $allLibraries += $result.Libraries
        }
        if ($result.Files) {
            $allFiles += $result.Files
        }
    }
    else {
        Write-Host "No data received or job encountered an error for job $($job.Id)"
    }
}

# Cleaning up the jobs
Get-Job | Remove-Job

Write-Host "Starting CSV files export"

# Exporting lists to CSV files
try {
    $allSites | Export-Csv -Path "${basePath}Sites_tenantTest.csv" -NoTypeInformation
    $allLibraries | Export-Csv -Path "${basePath}Libraries_tenantTest.csv" -NoTypeInformation
    $allFiles | Export-Csv -Path "${basePath}Files_tenantTest.csv" -NoTypeInformation
}
catch [System.IO.IOException] {
    Write-Host "File access error during export: $_"
}
catch {
    Write-Host "General error when exporting data to CSV: $_"
}

Write-Host "CSV files export finished"

$endTime = Get-Date
$endTimeStr = $endTime.ToString("dd-MM-yyyy HH:mm:ss")
Write-Host "Script end time: $endTimeStr"
$duration = $endTime - $startTime
$durationStr = "{0} hours, {1} minutes, {2} seconds" -f $duration.Hours, $duration.Minutes, $duration.Seconds
Write-Host "-------------------------- Script execution duration: $durationStr --------------------------"