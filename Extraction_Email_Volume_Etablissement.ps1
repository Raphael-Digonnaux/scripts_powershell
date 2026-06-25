#ATTENTION NECESSITE DE MODIFIER L'ADRESSE DE L'ADMIN (plus bas).

#requires -Modules ActiveDirectory

# Force TLS 1.2 pour Windows PowerShell 5.1
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

<#
.SYNOPSIS
    Exporte les boîtes Exchange Online avec leur volumétrie
    et le service provenant de l'Active Directory local.

.DESCRIPTION
    Le rapprochement entre Exchange Online et l'AD est effectué avec :
    1. l'adresse mail principale ;
    2. les attributs proxyAddresses ;
    3. l'attribut mail ;
    4. le UserPrincipalName.

    Le résultat est exporté dans un fichier CSV compatible Excel.
#>

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop

# ---------------------------------------------------------------------------
# PARAMÈTRES
# ---------------------------------------------------------------------------

$CheminExport = "C:\Temp\Volumetrie-Boites-O365.csv"
$CompteAdministrateurEXO = "monadresseadmin@mondomaine.fr"

# Types de boîtes à intégrer.
# Retirez "SharedMailbox" si vous voulez uniquement les boîtes nominatives.
$TypesBoites = @(
    "UserMailbox"
    "SharedMailbox"
)

# ---------------------------------------------------------------------------
# FONCTION DE CONVERSION DE LA TAILLE EN GO
# ---------------------------------------------------------------------------

function Convert-ToGigabytes {
    param (
        [Parameter(Mandatory = $false)]
        $TotalItemSize
    )

    if ($null -eq $TotalItemSize) {
        return 0
    }

    # Première méthode : objet Exchange complet
    try {
        if ($null -ne $TotalItemSize.Value) {
            $bytes = $TotalItemSize.Value.ToBytes()
            return [math]::Round($bytes / 1GB, 2)
        }
    }
    catch {
        # L'objet retourné par Exchange Online peut être désérialisé.
    }

    # Deuxième méthode : extraction du nombre d'octets indiqué entre parenthèses
    $texte = $TotalItemSize.ToString()

    # Exemples possibles :
    # 1.234 GB (1,324,567,890 bytes)
    # 1.234 GB (1 324 567 890 bytes)
    if ($texte -match '\(([\d\s,\.]+)\s+(bytes|octets)\)') {
        $nombreOctets = $matches[1] -replace '[^\d]', ''

        if ($nombreOctets) {
            return [math]::Round(([double]$nombreOctets / 1GB), 2)
        }
    }

    # Troisième méthode : conversion depuis la valeur affichée
    if ($texte -match '([\d\.,]+)\s*(B|KB|MB|GB|TB)') {
        $valeur = $matches[1] -replace ',', '.'
        $unite  = $matches[2].ToUpper()

        try {
            $nombre = [double]::Parse(
                $valeur,
                [System.Globalization.CultureInfo]::InvariantCulture
            )

            switch ($unite) {
                "B"  { return [math]::Round($nombre / 1GB, 2) }
                "KB" { return [math]::Round($nombre / 1MB, 2) }
                "MB" { return [math]::Round($nombre / 1KB, 2) }
                "GB" { return [math]::Round($nombre, 2) }
                "TB" { return [math]::Round($nombre * 1024, 2) }
            }
        }
        catch {
            return $null
        }
    }

    return $null
}

# ---------------------------------------------------------------------------
# PRÉPARATION DU DOSSIER D'EXPORT
# ---------------------------------------------------------------------------

$dossierExport = Split-Path -Path $CheminExport -Parent

if (-not (Test-Path -Path $dossierExport)) {
    New-Item -Path $dossierExport -ItemType Directory -Force | Out-Null
}

# ---------------------------------------------------------------------------
# CONNEXION EXCHANGE ONLINE
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Connexion à Exchange Online..." -ForegroundColor Cyan

try {
    # Ferme une éventuelle session Exchange Online résiduelle
    Disconnect-ExchangeOnline `
        -Confirm:$false `
        -ErrorAction SilentlyContinue

    Connect-ExchangeOnline `
        -UserPrincipalName $CompteAdministrateurEXO `
        -DisableWAM `
        -ShowBanner:$false `
        -ErrorAction Stop

    # Vérification effective de la connexion
    $ConnexionEXO = Get-ConnectionInformation -ErrorAction Stop

    if (-not $ConnexionEXO) {
        throw "La connexion Exchange Online n'a pas été établie."
    }

    Write-Host "Connexion Exchange Online établie." -ForegroundColor Green

    # -----------------------------------------------------------------------
    # RÉCUPÉRATION DE L'ACTIVE DIRECTORY
    # -----------------------------------------------------------------------

    Write-Host "Récupération des utilisateurs Active Directory..." -ForegroundColor Cyan

    $UtilisateursAD = Get-ADUser `
        -Filter * `
        -Properties GivenName,
                    Surname,
                    DisplayName,
                    mail,
                    UserPrincipalName,
                    proxyAddresses,
                    Department,
                    Enabled

    Write-Host "$($UtilisateursAD.Count) utilisateurs AD récupérés." -ForegroundColor Green

    # -----------------------------------------------------------------------
    # CRÉATION DE L'INDEX DES ADRESSES AD
    # -----------------------------------------------------------------------

    Write-Host "Création de l'index de rapprochement AD..." -ForegroundColor Cyan

    $IndexAD = @{}
    $AdressesDupliquees = @{}

    foreach ($UtilisateurAD in $UtilisateursAD) {

        $Adresses = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )

        if ($UtilisateurAD.mail) {
            [void]$Adresses.Add($UtilisateurAD.mail.Trim().ToLower())
        }

        if ($UtilisateurAD.UserPrincipalName) {
            [void]$Adresses.Add(
                $UtilisateurAD.UserPrincipalName.Trim().ToLower()
            )
        }

        foreach ($ProxyAddress in $UtilisateurAD.proxyAddresses) {
            if ($ProxyAddress -match '^(?i)smtp:(.+)$') {
                [void]$Adresses.Add($matches[1].Trim().ToLower())
            }
        }

        foreach ($Adresse in $Adresses) {
            if (-not $IndexAD.ContainsKey($Adresse)) {
                $IndexAD[$Adresse] = $UtilisateurAD
            }
            else {
                # Conservation de l'information en cas d'adresse présente
                # sur plusieurs comptes AD.
                if (-not $AdressesDupliquees.ContainsKey($Adresse)) {
                    $AdressesDupliquees[$Adresse] = @(
                        $IndexAD[$Adresse].SamAccountName
                    )
                }

                $AdressesDupliquees[$Adresse] += $UtilisateurAD.SamAccountName
            }
        }
    }

    Write-Host "$($IndexAD.Count) adresses AD indexées." -ForegroundColor Green

    if ($AdressesDupliquees.Count -gt 0) {
        Write-Warning "$($AdressesDupliquees.Count) adresse(s) sont présentes sur plusieurs comptes AD."
    }

    # -----------------------------------------------------------------------
    # RÉCUPÉRATION DES BOÎTES EXCHANGE ONLINE
    # -----------------------------------------------------------------------

    Write-Host "Récupération des boîtes Exchange Online..." -ForegroundColor Cyan

    $NombreTentativesMaximum = 4
    $Boites = $null
    $DerniereErreur = $null

    for (
        $Tentative = 1
        $Tentative -le $NombreTentativesMaximum
        $Tentative++
    ) {
        try {
            Write-Host (
                "Tentative {0}/{1}..." -f
                $Tentative,
                $NombreTentativesMaximum
            ) -ForegroundColor DarkCyan

            # Les propriétés nécessaires sont retournées par défaut.
            # Ne pas demander FirstName/LastName, non valides avec Get-EXOMailbox.
            $Boites = @(
                Get-EXOMailbox `
                    -ResultSize Unlimited `
                    -ErrorAction Stop |
                Where-Object {
                    $_.RecipientTypeDetails -in $TypesBoites
                } |
                Sort-Object DisplayName
            )

            if ($Boites.Count -eq 0) {
                throw "Exchange Online n'a retourné aucune boîte."
            }

            $DerniereErreur = $null
            break
        }
        catch {
            $DerniereErreur = $_

            Write-Warning (
                "Échec de la tentative {0} : {1}" -f
                $Tentative,
                $_.Exception.Message
            )

            if ($Tentative -lt $NombreTentativesMaximum) {
                $Attente = $Tentative * 10

                Write-Host (
                    "Nouvelle tentative dans {0} secondes..." -f $Attente
                ) -ForegroundColor Yellow

                Start-Sleep -Seconds $Attente
            }
        }
    }

    if ($null -ne $DerniereErreur -or $null -eq $Boites -or $Boites.Count -eq 0) {
        throw (
            "Impossible de récupérer les boîtes Exchange Online après " +
            "$NombreTentativesMaximum tentatives. " +
            "Dernière erreur : $($DerniereErreur.Exception.Message)"
        )
    }

    Write-Host "$($Boites.Count) boîtes Exchange Online récupérées." -ForegroundColor Green

    # -----------------------------------------------------------------------
    # ANALYSE DES BOÎTES
    # -----------------------------------------------------------------------

    $Resultats = [System.Collections.Generic.List[object]]::new()

    $Numero = 0

    foreach ($Boite in $Boites) {
        $Numero++

        $pourcentage = [math]::Round(
            ($Numero / $Boites.Count) * 100,
            0
        )

        Write-Progress `
            -Activity "Analyse des boîtes Exchange Online" `
            -Status "$Numero / $($Boites.Count) : $($Boite.DisplayName)" `
            -PercentComplete $pourcentage

        $AdressePrincipale = $Boite.PrimarySmtpAddress.ToString().ToLower()

        # -------------------------------------------------------------------
        # RAPPROCHEMENT AVEC L'AD
        # -------------------------------------------------------------------

        $UtilisateurAD = $null
        $MethodeCorrespondance = "Non trouvé"

        if ($IndexAD.ContainsKey($AdressePrincipale)) {
            $UtilisateurAD = $IndexAD[$AdressePrincipale]
            $MethodeCorrespondance = "Adresse SMTP principale"
        }
        elseif (
            $Boite.UserPrincipalName -and
            $IndexAD.ContainsKey($Boite.UserPrincipalName.ToLower())
        ) {
            $UtilisateurAD = $IndexAD[$Boite.UserPrincipalName.ToLower()]
            $MethodeCorrespondance = "UserPrincipalName"
        }

        # -------------------------------------------------------------------
        # RÉCUPÉRATION DES STATISTIQUES
        # -------------------------------------------------------------------

        $Statistiques = $null
        $ErreurStatistiques = $null

        try {
            $Statistiques = Get-EXOMailboxStatistics `
                -Identity $Boite.UserPrincipalName `
                -ErrorAction Stop
        }
        catch {
            $ErreurStatistiques = $_.Exception.Message
        }

        if ($Statistiques) {
            $TailleGo       = Convert-ToGigabytes $Statistiques.TotalItemSize
            $NombreElements = $Statistiques.ItemCount
            $DernierAcces   = $Statistiques.LastLogonTime
        }
        else {
            $TailleGo       = $null
            $NombreElements = $null
            $DernierAcces   = $null
        }

        # -------------------------------------------------------------------
        # CHOIX DU NOM ET DU PRÉNOM
        # Priorité donnée à l'AD pour les boîtes nominatives.
        # -------------------------------------------------------------------

        if ($UtilisateurAD) {
            $Nom       = $UtilisateurAD.Surname
            $Prenom    = $UtilisateurAD.GivenName
            $Service   = $UtilisateurAD.Department
            $CompteAD  = $UtilisateurAD.SamAccountName
            $CompteActifAD = $UtilisateurAD.Enabled
            $TrouveAD   = "Oui"
        }
        else {
            $Nom             = $null
            $Prenom          = $null
            $Service         = $null
            $CompteAD        = $null
            $CompteActifAD   = $null
            $TrouveAD        = "Non"
        }

        $AdresseDupliqueeAD = if (
            $AdressesDupliquees.ContainsKey($AdressePrincipale)
        ) {
            "Oui"
        }
        else {
            "Non"
        }

        # -------------------------------------------------------------------
        # AJOUT AU TABLEAU
        # -------------------------------------------------------------------

        $Resultats.Add(
            [PSCustomObject][ordered]@{
                Nom                         = $Nom
                Prenom                      = $Prenom
                NomAffiche                  = $Boite.DisplayName
                EmailPrincipal              = $Boite.PrimarySmtpAddress
                TailleOccupeeGo             = $TailleGo
                NombreElements              = $NombreElements
                ServiceAD                   = $Service
                CompteAD                    = $CompteAD
                CompteADActif               = $CompteActifAD
                TypeBoite                   = $Boite.RecipientTypeDetails
                CorrespondanceAD            = $TrouveAD
                MethodeCorrespondance        = $MethodeCorrespondance
                AdresseDupliqueeDansAD       = $AdresseDupliqueeAD
                DernierAccesBoite            = $DernierAcces
                ErreurRecuperationStatistiques = $ErreurStatistiques
            }
        )
    }

    Write-Progress `
        -Activity "Analyse des boîtes Exchange Online" `
        -Completed

    # -----------------------------------------------------------------------
    # EXPORT CSV
    # -----------------------------------------------------------------------

    if ($Resultats.Count -eq 0) {
        throw "Aucun résultat à exporter. Le fichier CSV ne sera pas créé."
    }

    $Resultats |
        Sort-Object ServiceAD, Nom, Prenom |
        Export-Csv `
            -Path $CheminExport `
            -Delimiter ";" `
            -NoTypeInformation `
            -Encoding UTF8

    # -----------------------------------------------------------------------
    # BILAN
    # -----------------------------------------------------------------------

    $NombreCorrespondances = @(
        $Resultats | Where-Object CorrespondanceAD -eq "Oui"
    ).Count

    $NombreSansCorrespondance = @(
        $Resultats | Where-Object CorrespondanceAD -eq "Non"
    ).Count

    $TailleTotale = (
        $Resultats |
        Measure-Object -Property TailleOccupeeGo -Sum
    ).Sum

    Write-Host ""
    Write-Host "Export terminé." -ForegroundColor Green
    Write-Host "Fichier : $CheminExport" -ForegroundColor Green
    Write-Host ""
    Write-Host "Boîtes analysées         : $($Resultats.Count)"
    Write-Host "Correspondances AD       : $NombreCorrespondances"
    Write-Host "Sans correspondance AD   : $NombreSansCorrespondance"
    Write-Host "Volumétrie totale        : $([math]::Round($TailleTotale, 2)) Go"
}
catch {
    Write-Host ""
    Write-Host "Le script s'est arrêté sur une erreur." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    throw
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}