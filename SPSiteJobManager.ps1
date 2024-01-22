# Définition des variables de base
$startTime = Get-Date
$startTimeStr = $startTime.ToString("dd-MM-yyyy HH:mm:ss")

$path = "C:\Users\ichennouf\OneDrive - ELIADIS\Desktop\scripts\Resultat\"
$cheminCsv = "${path}listeSites.csv"
$basePath = "${path}ExportCSV\"

# Limite pour le nombre de jobs parallèles
$jobLimit = 10

# Liste pour suivre les jobs
$runningJobs = @()

# Initialisation des listes pour stocker les informations
$allSites = @()
$allLibraries = @()
$allFolders = @()
$allFiles = @()

Write-Host "-------------------------- Début du script à $startTimeStr --------------------------"

# Importation du CSV
try {
    $listeSites = Import-Csv -Path $cheminCsv
    Write-Host "Importation du CSV réussie."
}
catch [System.IO.IOException] {
    Write-Host "Erreur d'accès au fichier CSV : $_"
    exit
}
catch [System.TimeoutException] {
    Write-Host "Le délai d'attente de l'opération a expiré : $_"
    exit
}
catch {
    Write-Host "Erreur générale lors de l'importation du CSV : $_"
    exit
}

# On demande les identifiants une seule fois
$cred = Get-Credential
if (-not $cred) {
    Write-Host "Aucun identifiant fourni. Le script va se terminer."
    exit
}

# On lance un job pour chaque site
foreach ($site in $listeSites) {

    # On lance un job si le nombre de jobs en cours est inférieur à la limite
    while ($runningJobs.Count -ge $jobLimit) {
        # On attend qu'un job se termine
        $completedJob = Wait-Job -Job $runningJobs -Any
        $runningJobs = $runningJobs | Where-Object { $_.Id -ne $completedJob.Id }
    }

    Write-Host "Début du traitement du site : $($site.Url)"

    # On lance un job par site
    $job = Start-Job -ScriptBlock {
        function Get-FolderSize {
            param (
                [Parameter(Mandatory = $true)]
                $folderFiles
            )
            # Logique pour calculer la taille du dossier
            $totalSize = 0
            foreach ($file in $folderFiles) {
                if ($file.FileSystemObjectType -eq "File") {
                    $totalSize += $file.FieldValues.File_x0020_Size
                }
            }
            return $totalSize
        }

        # Fonction pour traiter chaque site SharePoint
        function ProcessSite {
            param (
                [Parameter(Mandatory = $true)]
                [PSCustomObject]$site,

                [Parameter(Mandatory = $true)]
                [System.Management.Automation.PSCredential]$cred
            )

            $pageSize = 1000

            # Initialisation des listes locales
            $localSites = @()
            $localLibraries = @()
            $localFolders = @()
            $localFiles = @()

            try {
                try {
                    # Connexion à chaque site avec les identifiants
                    $connectionSite = Connect-PnPOnline -Url $site.Url -Credentials $cred -ReturnConnection

                    # On récupère les informations du site
                    $web = Get-PnPWeb -Connection $connectionSite

                    # Créer un objet personnalisé pour les informations du site
                    $siteInfo = [PSCustomObject]@{
                        Title   = $web.Title
                        SiteUrl = $web.Url
                        Owner   = $web.Owner
                    }
                    $localSites += $siteInfo
                }
                catch [System.Net.WebException] {
                    Write-Host "Erreur de réseau lors de la connexion à SharePoint : $_"
                }
                catch [Microsoft.SharePoint.Client.IdcrlException] {
                    Write-Host "Erreur d'authentification SharePoint : $_"
                }
                catch {
                    Write-Host "Erreur inconnue lors de la connexion à SharePoint : $_"
                }

                try {
                    # On récupère toutes les bibliothèques de documents du site
                    $bibliotheques = Get-PnPList -Connection $connectionSite | Where-Object { $_.BaseTemplate -eq 101 }

                    foreach ($bibliotheque in $bibliotheques) {
                        # On ajoute les informations de la bibliothèque à la liste $allLibraries
                        $libraryInfo = [PSCustomObject]@{
                            SiteUrl      = $site.Url
                            LibraryTitle = $bibliotheque.Title
                            TotalSize    = 0  # Initialisation de la taille totale de la bibliothèque
                            FileCount    = 0  # Initialisation du nombre de fichiers dans la bibliothèque
                        }
                        $localLibraries += $libraryInfo

                        try {
                            # On récupère tous les fichiers et dossiers de chaque bibliothèque
                            $listItems = Get-PnPListItem -List $bibliotheque -PageSize $pageSize -Connection $ConnectionSite

                            foreach ($item in $listItems) {
                                if ($item.FileSystemObjectType -eq "File") {
                                    # Si l'élément est un fichier
                                    $fileInfo = [PSCustomObject]@{
                                        SiteUrl       = $site.Url
                                        Library       = $bibliotheque.Title
                                        FileName      = $item.FieldValues.FileLeafRef
                                        FilePath      = $item.FieldValues.FileDirRef
                                        FileSize      = $item.FieldValues.File_x0020_Size
                                        CreatedDate   = $item.FieldValues.Created
                                        ModifiedDate  = $item.FieldValues.Modified
                                        FileExtension = [System.IO.Path]::GetExtension($item.FieldValues.FileLeafRef)
                                    }
                                    $localFiles += $fileInfo

                                    # On ajoute la taille du fichier à la taille totale de la bibliothèque
                                    $libraryInfo.TotalSize += $item.FieldValues.File_x0020_Size

                                    # On met à jour le nombre de fichiers dans la bibliothèque
                                    $libraryInfo.FileCount++
                                }
                                elseif ($item.FileSystemObjectType -eq "Folder") {
                                    try {
                                        # Si l'élément est un dossier
                                        $folderFiles = Get-PnPListItem -List $bibliotheque -Folder $item.FieldValues.FileDirRef -Connection $connectionSite

                                        # On calcule la taille du dossier
                                        $folderSize = Get-FolderSize -folderFiles $folderFiles

                                        # On ajoute les informations du dossier à la liste $allFolders
                                        $folderInfo = [PSCustomObject]@{
                                            SiteUrl      = $site.Url
                                            Library      = $bibliotheque.Title
                                            FolderName   = $item.FieldValues.FileLeafRef
                                            FolderPath   = $item.FieldValues.FileDirRef
                                            FolderSize   = $folderSize
                                            FolderLength = ($folderFiles | Where-Object { $_.FileSystemObjectType -eq "File" }).Count + $folderFiles.Count
                                        }
                                        $localFolders += $folderInfo
                                    }
                                    catch {
                                        Write-Host "Erreur lors du traitement des dossiers du site $($site.Url) : $_"
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Host "Erreur lors de la récupération des fichiers et dossiers du site $($site.Url) : $_"
                        }
                    }
                }
                catch {
                    Write-Host "Erreur lors de la récupération des bibliothèques de documents du site $($site.Url) : $_"
                }
            }
            catch [Microsoft.SharePoint.Client.ClientRequestException] {
                Write-Host "Erreur de requête SharePoint du site $($site.Url) : $_"
            }
            catch [System.Net.WebException] {
                Write-Host "Erreur de réseau lors de la connexion à SharePoint du site $($site.Url) : $_"
            }
            catch {
                Write-Host "Erreur inconnue lors de la connexion ou de la récupération des données du site $($site.Url) : $_"
            }

            # On retourne un objet contenant toutes les listes
            return @{
                Sites     = $localSites
                Libraries = $localLibraries
                Files     = $localFiles
                Folders   = $localFolders
            }
        }

        # On exécute la fonction ProcessSite
        ProcessSite -site $args[0] -cred $args[1]
    } -ArgumentList $site, $cred
    $runningJobs += $job
}

Write-Host "Attente de la fin de tous les jobs..."
# On attend la fin de tous les jobs
Get-Job | Wait-Job
Write-Host "Tous les jobs sont terminés."

# On récupère les résultats de chaque job
$jobs = Get-Job

# Traitement des jobs
foreach ($job in $jobs) {

    Write-Host "Traitement des résultats du job $($job.Id)"
    $result = Receive-Job -Job $job

    if ($job.State -eq "Failed") {
        Write-Host "Le job $($job.Id) a échoué."
        continue
    }

    # On vérifie que le résultat n'est pas null
    if ($result) {
        # On ajoute les résultats aux listes globales si elles sont présentes, indépendamment les unes des autres
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
        Write-Host "Aucune donnée reçue ou job a rencontré une erreur pour le job $($job.Id)."
    }
}


# On nettoie les jobs
Get-Job | Remove-Job

Write-Host "Début d'exportation des fichiers CSV"

# On exporte les listes en fichiers CSV
try {
    $allSites | Export-Csv -Path "${basePath}Sites_tenantTest.csv" -NoTypeInformation
    $allLibraries | Export-Csv -Path "${basePath}Libraries_tenantTest.csv" -NoTypeInformation
    $allFiles | Export-Csv -Path "${basePath}Files_tenantTest.csv" -NoTypeInformation
    $allFolders | Export-Csv -Path "${basePath}Folders_tenantTest.csv" -NoTypeInformation
}
catch [System.IO.IOException] {
    Write-Host "Erreur d'accès au fichier lors de l'exportation : $_"
}
catch {
    Write-Host "Erreur générale lors de l'exportation des données vers CSV : $_"
}

Write-Host "Fin d'exportation des fichiers CSV"

$endTime = Get-Date
$endTimeStr = $endTime.ToString("dd-MM-yyyy HH:mm:ss")
Write-Host "Heure de fin du script : $endTimeStr"
$duration = $endTime - $startTime
$durationStr = "{0} heures, {1} minutes, {2} secondes" -f $duration.Hours, $duration.Minutes, $duration.Seconds
Write-Host "-------------------------- Durée d'exécution du script : $durationStr --------------------------"