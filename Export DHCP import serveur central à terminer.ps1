#Test emplacement où est le fichier
if ($MyInvocation.MyCommand.Path) {
    $scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Path
} else {
    $scriptDirectory = $PWD.Path
}

#On place le csv des serveurs au même endroit que le script
#le csv doit ressembler à ceci sans les #
#ServerName
#serv-mdo-dc
#serv-other-dc
#serv-another-dc

# Chemin du fichier CSV contenant les noms des serveurs
$csvPath = "$scriptDirectory\servers.csv"

# Chemin du fichier de sortie pour chaque serveur DHCP exporté
$tempExportPath = "$scriptDirectory\TempDHCPExport"

# Serveur central où importer les configurations DHCP
$centralServer = "serv-dc02"

# Lire la liste des serveurs depuis le fichier CSV
$servers = Import-Csv -Path $csvPath

# Créer un dossier temporaire si nécessaire
if (-not (Test-Path -Path $tempExportPath)) {
    New-Item -Path $tempExportPath -ItemType Directory
}

# Boucle à travers chaque serveur de la liste pour exporter et importer les données DHCP
foreach ($server in $servers) {
    $serverName = $server.ServerName

    # Définir le chemin du fichier d'exportation temporaire pour ce serveur
    $exportFilePath = Join-Path -Path $tempExportPath -ChildPath "$serverName.xml"

    # Exporter les données DHCP du serveur vers un fichier XML
    Write-Host "Exporting DHCP configuration from $serverName..."
    Export-DhcpServer -ComputerName $serverName -Leases -File $exportFilePath

    # Vérifier si l'exportation a réussi
    if (Test-Path $exportFilePath) {
        Write-Host "Export successful for $serverName. Importing into $centralServer..."

        # Importer les données DHCP dans le serveur central
        Import-DhcpServer -ComputerName $centralServer -Leases -File $exportFilePath -BackupPath "C:\Users\admin.rd\Desktop\backup" -ScopeOverwrite

        Write-Host "Import successful for $serverName into $centralServer."

        # Supprimer le fichier exporté après l'importation
        Remove-Item -Path $exportFilePath
    } else {
        Write-Host "Failed to export DHCP configuration from $serverName."
    }
}

# Nettoyer les fichiers temporaires
Write-Host "Cleaning up temporary export directory..."
Remove-Item -Path $tempExportPath -Recurse -Force
Write-Host "Process completed."
