# Interactive connection to SharePoint service to obtain the list of sites
$SiteAdminURL = "https://m365x52471717-admin.sharepoint.com"
$adminConnection = Connect-PnPOnline -Url $SiteAdminURL -Interactive -ReturnConnection

# Retrieve the list of all SharePoint Online sites
$sitesList = Get-PnPTenantSite -Connection $adminConnection

# Export the URLs of the sites to a CSV file
$sitesList | Select-Object Url | Export-Csv -Path "C:\Users\ichennouf\OneDrive - ELIADIS\Desktop\scripts\Resultat\listSites.csv" -NoTypeInformation
