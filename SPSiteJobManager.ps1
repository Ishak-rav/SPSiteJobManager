# Définition des variables de base
$startTime = Get-Date
$path = "C:\Users\ichennouf\OneDrive - ELIADIS\Desktop\scripts\Resultat\"
$cheminCsv = "${path}listeSites.csv"
$basePath = "${path}ExportCSV\"
$logPath = "${basePath}logs\scriptLog.log"

# Fonction pour écrire dans le fichier de logs
function Write-Log {
    Param (
        [string]$message,
        [bool]$isError = $false
    )
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timeStamp : $message"
    $logMessage | Out-File -FilePath $logPath -Append

    if ($isError) {
        Write-Error $message
    }
}

Write-Log "-------------------------- Début du script à $startTime --------------------------"

# Importation du CSV
try {
    $listeSites = Import-Csv -Path $cheminCsv
    Write-Log "Importation du CSV réussie."
}
catch [System.IO.IOException] {
    Write-Log "Erreur d'accès au fichier CSV : $_" -isError $true
    exit
}
catch [System.TimeoutException] {
    Write-Log "Le délai d'attente de l'opération a expiré : $_" -isError $true
    exit
}
catch {
    Write-Log "Erreur générale lors de l'importation du CSV : $_" -isError $true
    exit
}

# Demander les identifiants une seule fois
$cred = Get-Credential
if (-not $cred) {
    Write-Log "Aucun identifiant fourni. Le script va se terminer." -isError $true
    exit
}

# Initialiser les listes pour stocker les informations
$allSites = @()
$allLibraries = @()
$allFolders = @()
$allFiles = @()

# Lancer un job pour chaque site
foreach ($site in $listeSites) {

    Write-Log "Début du traitement du site : $($site.Url)"

    Start-Job -ScriptBlock {
        function Write-Log {
            Param (
                [Parameter(Mandatory = $true)]
                [string]$message,

                [bool]$isError = $false
            )
            $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "$timeStamp : $message"
            $logMessage | Out-File -FilePath $logPath -Append

            if ($isError) {
                Write-Error $message
            }
        }

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

            # Initialisation des listes locales
            $localSites = @()
            $localLibraries = @()
            $localFolders = @()
            $localFiles = @()

            try {
                try {
                    # Connexion à chaque site avec les identifiants
                    $connectionSite = Connect-PnPOnline -Url $site.Url -Credentials $cred -ReturnConnection

                    # Récupérer les informations du site
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
                    Write-Log "Erreur de réseau lors de la connexion à SharePoint : $_" -isError $true
                }
                catch [Microsoft.SharePoint.Client.IdcrlException] {
                    Write-Log "Erreur d'authentification SharePoint : $_" -isError $true
                }
                catch {
                    Write-Log "Erreur inconnue lors de la connexion à SharePoint : $_" -isError $true
                }

                try {
                    # Récupérer toutes les bibliothèques de documents du site
                    $bibliotheques = Get-PnPList -Connection $connectionSite | Where-Object { $_.BaseTemplate -eq 101 }

                    foreach ($bibliotheque in $bibliotheques) {
                        # Ajouter les informations de la bibliothèque à la liste $allLibraries
                        $libraryInfo = [PSCustomObject]@{
                            SiteUrl      = $site.Url
                            LibraryTitle = $bibliotheque.Title
                            TotalSize    = 0  # Initialisation de la taille totale de la bibliothèque
                            FileCount    = 0  # Initialisation du nombre de fichiers dans la bibliothèque
                        }
                        $localLibraries += $libraryInfo

                        try {

                            # Récupérer tous les fichiers et dossiers de chaque bibliothèque
                            $items = Get-PnPListItem -List $bibliotheque -Connection $connectionSite

                            foreach ($item in $items) {
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

                                    # Ajouter la taille du fichier à la taille totale de la bibliothèque
                                    $libraryInfo.TotalSize += $item.FieldValues.File_x0020_Size

                                    # Mettre à jour le nombre de fichiers dans la bibliothèque
                                    $libraryInfo.FileCount++
                                }
                                elseif ($item.FileSystemObjectType -eq "Folder") {
                                    try {
                                        # Si l'élément est un dossier
                                        $folderFiles = Get-PnPListItem -List $bibliotheque -Folder $item.FieldValues.FileDirRef -Connection $connectionSite

                                        # Calculer la taille du dossier
                                        $folderSize = Get-FolderSize -folderFiles $folderFiles

                                        # Ajouter les informations du dossier à la liste $allFolders
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
                                        Write-Log "Erreur lors du traitement des dossiers : $_" -isError $true
                                    }
                                }
                            }
                        }
                        catch {
                            Write-Log "Erreur lors de la récupération des fichiers et dossiers : $_" -isError $true
                        }
                    }
                }
                catch {
                    Write-Log "Erreur lors de la récupération des bibliothèques de documents : $_" -isError $true
                }
            }
            catch [Microsoft.SharePoint.Client.ClientRequestException] {
                Write-Log "Erreur de requête SharePoint : $_" -isError $true
            }
            catch [System.Net.WebException] {
                Write-Log "Erreur de réseau lors de la connexion à SharePoint : $_" -isError $true
            }
            catch {
                Write-Log "Erreur inconnue lors de la connexion ou de la récupération des données : $_" -isError $true
            }

            # Retourner un objet contenant toutes les listes
            return @{
                Sites     = $localSites
                Libraries = $localLibraries
                Files     = $localFiles
                Folders   = $localFolders
            }
        }

        # Exécution de la fonction ProcessSite
        ProcessSite -site $args[0] -cred $args[1]
    } -ArgumentList $site, $cred, $logPath
}

Write-Log "Attente de la fin de tous les jobs..."
# Attendre la fin de tous les jobs
Get-Job | Wait-Job
Write-Log "Tous les jobs sont terminés."

# Récupérer les résultats de chaque job
$jobs = Get-Job

# Traitement des jobs
foreach ($job in $jobs) {

    Write-Log "Traitement des résultats du job $($job.Id)"
    $result = Receive-Job -Job $job

    if ($job.State -eq "Failed") {
        Write-Log "Le job $($job.Id) a échoué." -isError $true
        continue
    }

    # Vérifier que le résultat n'est pas null
    if ($result) {
        # Ajouter les résultats aux listes globales si elles sont présentes, indépendamment les unes des autres
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
        Write-Log "Aucune donnée reçue ou job a rencontré une erreur pour le job $($job.Id)." -isError $true
    }
}


# Nettoyer les jobs
Get-Job | Remove-Job

Write-Log "Début d'exportation des fichiers CSV"

# Exporter les listes en fichiers CSV
try {
    $allSites | Export-Csv -Path "${basePath}Sites_tenantTest.csv" -NoTypeInformation
    $allLibraries | Export-Csv -Path "${basePath}Libraries_tenantTest.csv" -NoTypeInformation
    $allFiles | Export-Csv -Path "${basePath}Files_tenantTest.csv" -NoTypeInformation
    $allFolders | Export-Csv -Path "${basePath}Folders_tenantTest.csv" -NoTypeInformation
}
catch [System.IO.IOException] {
    Write-Log "Erreur d'accès au fichier lors de l'exportation : $_" -isError $true
}
catch {
    Write-Log "Erreur générale lors de l'exportation des données vers CSV : $_" -isError $true
}

Write-Log "Fin d'exportation des fichiers CSV"

$endTime = Get-Date
Write-Log "Heure de fin du script : $endTime"
$duration = $endTime - $startTime
Write-Log "-------------------------- Durée d'exécution du script : $($duration.Hours) heures, $($duration.Minutes) minutes, $($duration.Seconds) secondes --------------------------"