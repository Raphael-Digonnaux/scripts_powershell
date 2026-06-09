Import-Module ActiveDirectory
Import-Module ImportExcel

$Racine = "\\SERV-FILES\Mon_dossier"
$Export = "C:\Temp\Comptes_avec_MP3.xlsx"

# Récupération de la liste des dossiers personnels
$Dossiers = Get-ChildItem -Path $Racine -Directory -ErrorAction Stop

$Total = $Dossiers.Count
$Index = 0
$Resultats = @()

foreach ($Dossier in $Dossiers) {

    $Index++
    $Compte = $Dossier.Name
    $DossierPersonnel = $Dossier.FullName
    $Pourcentage = [math]::Round(($Index / $Total) * 100, 0)

    Write-Progress `
        -Activity "Recherche des fichiers MP3 dans les dossiers personnels" `
        -Status "Traitement du compte $Compte ($Index / $Total)" `
        -PercentComplete $Pourcentage

    Write-Host "[$Index/$Total] Analyse du compte : $Compte"

    $Mp3 = Get-ChildItem -Path $DossierPersonnel -Recurse -File -Filter "*.mp3" -ErrorAction SilentlyContinue

    if ($Mp3) {

        $UserAD = Get-ADUser `
            -Filter "SamAccountName -eq '$Compte'" `
            -Properties Department, DisplayName, Enabled `
            -ErrorAction SilentlyContinue

        $Resultats += [PSCustomObject]@{
            Compte          = $Compte
            NomComplet      = $UserAD.DisplayName
            Service         = $UserAD.Department
            CompteActif     = $UserAD.Enabled
            Dossier         = $DossierPersonnel
            NombreMP3       = $Mp3.Count
            TailleTotaleMo  = [math]::Round(($Mp3 | Measure-Object Length -Sum).Sum / 1MB, 2)
        }

        Write-Host "    -> MP3 trouvés : $($Mp3.Count)" -ForegroundColor Yellow
    }
}

Write-Progress `
    -Activity "Recherche des fichiers MP3 dans les dossiers personnels" `
    -Completed

$Resultats |
    Sort-Object Service, Compte |
    Export-Excel -Path $Export `
        -WorksheetName "Comptes avec MP3" `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow

Write-Host ""
Write-Host "Analyse terminée." -ForegroundColor Green
Write-Host "Comptes avec MP3 : $($Resultats.Count)" -ForegroundColor Green
Write-Host "Export Excel : $Export" -ForegroundColor Green