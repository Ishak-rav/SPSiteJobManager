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
    catch {
        Write-Host "Unknown error connecting or retrieving site data $($site.Url) : $_"
    }

    # Returning an object containing all the lists
    return @{
        Sites     = $localSites
        Libraries = $localLibraries
        Files     = $localFiles
    }
}