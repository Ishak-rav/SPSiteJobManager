# Interactive connection to SharePoint service to obtain the list of sites
$SiteAdminURL = "<YourSharePointAdminURL>"
$adminConnection = Connect-PnPOnline -Url $SiteAdminURL -Interactive -ReturnConnection

# Retrieve the list of all SharePoint Online sites
$sitesList = Get-PnPTenantSite -Connection $adminConnection

# Export the URLs of the sites to a CSV file
$sitesList | Select-Object Url | Export-Csv -Path "<PathToYourScripts>\listSites.csv" -NoTypeInformation
