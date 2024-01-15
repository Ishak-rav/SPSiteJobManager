# Connexion interactive au service SharePoint pour obtenir la liste des sites
$SiteAdminURL = "https://m365x52471717-admin.sharepoint.com"
$adminConnection = Connect-PnPOnline -Url $SiteAdminURL -Interactive -ReturnConnection

# Récupérer la liste de tous les sites SharePoint Online
$listeSites = Get-PnPTenantSite -Connection $adminConnection

# Exporter les URL des sites dans un fichier CSV
$listeSites | Select-Object Url | Export-Csv -Path "C:\Users\ichennouf\OneDrive - ELIADIS\Desktop\scripts\Resultat\listeSites.csv" -NoTypeInformation
