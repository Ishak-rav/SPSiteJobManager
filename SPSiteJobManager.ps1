# Base variables definition
$moduleName = "PnP.PowerShell"
$startTime = Get-Date
$startTimeStr = $startTime.ToString("dd-MM-yyyy HH:mm:ss")

$path = "<YourScriptPath>\Result\"
$csvPath = "${path}sitesList.csv"
$basePath = "${path}ExportCSV\"

# Limit for the number of parallel jobs
$jobLimit = 10

# List to track the running jobs
$runningJobs = @()

# Initializing lists to store information
$allSites = @()
$allLibraries = @()
$allFolders = @()
$allFiles = @()

Write-Host "-------------------------- Script start at $startTimeStr --------------------------"

# Function to check if the module is available
function ModuleChecker {
    param (
        [string]$ModuleName
    )
    return Get-Module -ListAvailable -Name $ModuleName
}

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
$cred = Get-Credential
if (-not $cred) {
    Write-Host "No identifier provided. The script will end."
    exit
}

# Starting a job for each site
foreach ($site in $listeSites) {

    # Launch a job if the number of jobs in progress is less than the limit
    while ($runningJobs.Count -ge $jobLimit) {
        # Waiting for a job to finish
        $completedJob = Wait-Job -Job $runningJobs -Any
        $runningJobs = $runningJobs | Where-Object { $_.Id -ne $completedJob.Id }
    }

    Write-Host "Start of site processing : $($site.Url)"

    $job = Start-Job -ScriptBlock {
        function Get-FolderSize {
            param (
                [Parameter(Mandatory = $true)]
                $folderFiles
            )
            # Logic to calculate the folder size
            $totalSize = 0
            foreach ($file in $folderFiles) {
                if ($file.FileSystemObjectType -eq "File") {
                    $totalSize += $file.FieldValues.File_x0020_Size
                }
            }
            return $totalSize
        }

        # Function to process each SharePoint site
        function ProcessSite {
            param (
                [Parameter(Mandatory = $true)]
                [PSCustomObject]$site,

                [Parameter(Mandatory = $true)]
                [System.Management.Automation.PSCredential]$cred
            )

            $pageSize = 1000

            # Initializing local lists
            $localSites = @()
            $localLibraries = @()
            $localFolders = @()
            $localFiles = @()

            try {
                try {
                    # Connection to each site
                    $siteConnection = Connect-PnPOnline -Url $site.Url -Credentials $cred -ReturnConnection

                    # Retrieving the information from the site
                    $web = Get-PnPWeb -Connection $siteConnection

                    # Creating custom object for site information
                    $siteInfo = [PSCustomObject]@{
                        Title   = $web.Title
                        SiteUrl = $web.Url
                        Owner   = $web.Owner
                    }
                    $localSites += $siteInfo
                }
                catch [System.Net.WebException] {
                    Write-Host "Network error connecting to SharePoint: $_"
                }
                catch [Microsoft.SharePoint.Client.IdcrlException] {
                    Write-Host "SharePoint authentication error: $_"
                }
                catch {
                    Write-Host "Unknown error connecting to SharePoint: $_"
                }

                try {
                    # Retrieving all the document libraries on the site
                    $libraries = Get-PnPList -Connection $siteConnection | Where-Object { $_.BaseTemplate -eq 101 }

                    foreach ($library in $libraries) {
                        # Adding the library information to the $allLibraries list
                        $libraryInfo = [PSCustomObject]@{
                            SiteUrl      = $site.Url
                            LibraryTitle = $library.Title
                            TotalSize    = 0  # Initializing the total library size
                            FileCount    = 0  # Initializing the number of files in the library
                        }
                        $localLibraries += $libraryInfo

                        try {
                            # Recovering all the files and folders from each library
                            $listItems = Get-PnPListItem -List $library -PageSize $pageSize -Connection $siteConnection

                            foreach ($item in $listItems) {
                                if ($item.FileSystemObjectType -eq "File") {
                                    # If the item is a file
                                    $fileInfo = [PSCustomObject]@{
                                        SiteUrl       = $site.Url
                                        Library       = $library.Title
                                        FileName      = $item.FieldValues.FileLeafRef
                                        FilePath      = $item.FieldValues.FileDirRef
                                        FileSize      = $item.FieldValues.File_x0020_Size
                                        CreatedDate   = $item.FieldValues.Created
                                        ModifiedDate  = $item.FieldValues.Modified
                                        FileExtension = [System.IO.Path]::GetExtension($item.FieldValues.FileLeafRef)
                                    }
                                    $localFiles += $fileInfo

                                    # Adding the file size to the total size of the library
                                    $libraryInfo.TotalSize += $item.FieldValues.File_x0020_Size

                                    # Updating the number of files in the library
                                    $libraryInfo.FileCount++
                                }
                                elseif ($item.FileSystemObjectType -eq "Folder") {
                                    try {
                                        # If the item is a folder
                                        $folderFiles = Get-PnPListItem -List $library -Folder $item.FieldValues.FileDirRef -Connection $siteConnection

                                        # Calculating the size of the file
                                        $folderSize = Get-FolderSize -folderFiles $folderFiles

                                        # Adding the folder information to the $allFolders list
                                        $folderInfo = [PSCustomObject]@{
                                            SiteUrl      = $site.Url
                                            Library      = $library.Title
                                            FolderName   = $item.FieldValues.FileLeafRef
                                            FolderPath   = $item.FieldValues.FileDirRef
                                            FolderSize   = $folderSize
                                            FolderLength = ($folderFiles | Where-Object { $_.FileSystemObjectType -eq "File" }).Count + $folderFiles.Count
                                        }
                                        $localFolders += $folderInfo
                                    }
                                    catch {
                                        Write-Host "Error processing site folders $($site.Url) : $_"
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Host "Error retrieving files and folders from site $($site.Url) : $_"
                        }
                    }
                }
                catch {
                    Write-Host "Error retrieving document libraries from site $($site.Url) : $_"
                }
            }
            catch [Microsoft.SharePoint.Client.ClientRequestException] {
                Write-Host "SharePoint site query error $($site.Url) : $_"
            }
            catch [System.Net.WebException] {
                Write-Host "Network error connecting to SharePoint site $($site.Url) : $_"
            }
            catch {
                Write-Host "Unknown error connecting or retrieving site data $($site.Url) : $_"
            }

            # Returning an object containing all the lists
            return @{
                Sites     = $localSites
                Libraries = $localLibraries
                Files     = $localFiles
                Folders   = $localFolders
            }
        }

        # Executing the ProcessSite function
        ProcessSite -site $args[0] -cred $args[1]
    } -ArgumentList $site, $cred
    $runningJobs += $job
}

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
        if ($result.Folders) {
            $allFolders += $result.Folders
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
    $allSites | Export-Csv -Path "${basePath}SitesExport.csv" -NoTypeInformation
    $allLibraries | Export-Csv -Path "${basePath}LibrariesExport.csv" -NoTypeInformation
    $allFiles | Export-Csv -Path "${basePath}FilesExport.csv" -NoTypeInformation
    $allFolders | Export-Csv -Path "${basePath}FoldersExport.csv" -NoTypeInformation
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