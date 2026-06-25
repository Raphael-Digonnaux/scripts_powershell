Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# PARAMÈTRES
# ---------------------------------------------------------------------------

# Mettre à $true pendant les tests pour voir Excel.
# Une fois le fonctionnement validé, vous pourrez passer à $false.
$AfficherExcel = $true

# ---------------------------------------------------------------------------
# VÉRIFICATION DU MODE STA
# ---------------------------------------------------------------------------

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {

    [System.Windows.Forms.MessageBox]::Show(
        "Ce script doit être exécuté en mode STA.`n`n" +
        "Commande recommandée :`n" +
        "PowerShell.exe -STA -ExecutionPolicy Bypass -File ""C:\SCRIPTS\Separer-Onglets-Excel.ps1""",
        "Mode STA requis",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    exit
}

# ---------------------------------------------------------------------------
# FONCTIONS
# ---------------------------------------------------------------------------

function Select-ExcelFile {

    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    $dialog.Title = "Sélectionnez le classeur Excel à séparer"
    $dialog.Filter = "Classeurs Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Tous les fichiers (*.*)|*.*"
    $dialog.FilterIndex = 1
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $dialog.CheckPathExists = $true
    $dialog.RestoreDirectory = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }

    return $null
}

function Select-DestinationFolder {

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    $dialog.Description = "Sélectionnez le dossier de destination"
    $dialog.ShowNewFolderButton = $true

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

function Get-SafeFileName {

    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $nomNettoye = $Name

    foreach ($caractere in [System.IO.Path]::GetInvalidFileNameChars()) {
        $nomNettoye = $nomNettoye.Replace($caractere, "_")
    }

    $nomNettoye = $nomNettoye.Trim().TrimEnd(".")

    if ([string]::IsNullOrWhiteSpace($nomNettoye)) {
        $nomNettoye = "Onglet"
    }

    return $nomNettoye
}

function Get-UniqueFilePath {

    param(
        [Parameter(Mandatory = $true)]
        [string]$Folder,

        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $chemin = Join-Path -Path $Folder -ChildPath "$FileName$Extension"
    $compteur = 1

    while (Test-Path -LiteralPath $chemin) {

        $chemin = Join-Path `
            -Path $Folder `
            -ChildPath "$FileName ($compteur)$Extension"

        $compteur++
    }

    return $chemin
}

function Release-ComObject {

    param(
        [Parameter(Mandatory = $false)]
        $ComObject
    )

    if ($null -ne $ComObject) {

        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $ComObject
            ) | Out-Null
        }
        catch {
        }
    }
}

function Update-ProgressWindow {

    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $false)]
        [int]$Value = -1,

        [Parameter(Mandatory = $false)]
        [string]$Details = ""
    )

    $label.Text = $Text
    $labelDetails.Text = $Details

    if (
        $Value -ge $progressBar.Minimum -and
        $Value -le $progressBar.Maximum
    ) {
        $progressBar.Value = $Value
    }

    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------
# SÉLECTION DU FICHIER SOURCE
# ---------------------------------------------------------------------------

$FichierSource = Select-ExcelFile

if ([string]::IsNullOrWhiteSpace($FichierSource)) {

    [System.Windows.Forms.MessageBox]::Show(
        "Aucun fichier Excel n'a été sélectionné.",
        "Opération annulée",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    exit
}

if (-not (Test-Path -LiteralPath $FichierSource)) {

    [System.Windows.Forms.MessageBox]::Show(
        "Le fichier sélectionné n'existe pas.",
        "Fichier introuvable",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit
}

# ---------------------------------------------------------------------------
# SÉLECTION DU DOSSIER DE DESTINATION
# ---------------------------------------------------------------------------

$DossierDestination = Select-DestinationFolder

if ([string]::IsNullOrWhiteSpace($DossierDestination)) {

    [System.Windows.Forms.MessageBox]::Show(
        "Aucun dossier de destination n'a été sélectionné.",
        "Opération annulée",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )

    exit
}

if (-not (Test-Path -LiteralPath $DossierDestination)) {

    [System.Windows.Forms.MessageBox]::Show(
        "Le dossier de destination n'existe pas ou n'est pas accessible.",
        "Dossier inaccessible",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    exit
}

# ---------------------------------------------------------------------------
# FENÊTRE DE PROGRESSION
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form

$form.Text = "Séparation des onglets Excel"
$form.Size = New-Object System.Drawing.Size(650, 210)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.ControlBox = $false
$form.TopMost = $true

$label = New-Object System.Windows.Forms.Label

$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(590, 50)
$label.Text = "Initialisation..."
$label.AutoEllipsis = $true

$form.Controls.Add($label)

$progressBar = New-Object System.Windows.Forms.ProgressBar

$progressBar.Location = New-Object System.Drawing.Point(20, 80)
$progressBar.Size = New-Object System.Drawing.Size(590, 25)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

$form.Controls.Add($progressBar)

$labelDetails = New-Object System.Windows.Forms.Label

$labelDetails.Location = New-Object System.Drawing.Point(20, 120)
$labelDetails.Size = New-Object System.Drawing.Size(590, 40)
$labelDetails.Text = ""
$labelDetails.AutoEllipsis = $true

$form.Controls.Add($labelDetails)

# ---------------------------------------------------------------------------
# VARIABLES
# ---------------------------------------------------------------------------

$excel = $null
$workbooks = $null
$classeurSource = $null
$worksheets = $null
$onglet = $null
$nouveauClasseur = $null

$nombreFichiersCrees = 0
$erreurs = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------------------
# TRAITEMENT
# ---------------------------------------------------------------------------

try {

    $form.Show()

    Update-ProgressWindow `
        -Text "Démarrage de Microsoft Excel..." `
        -Details "Veuillez patienter."

    $excel = New-Object -ComObject Excel.Application

    $excel.Visible = $AfficherExcel
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AlertBeforeOverwriting = $false

    try {
        # Désactivation des macros
        $excel.AutomationSecurity = 3
    }
    catch {
    }

    try {
        # Calcul manuel
        $excel.Calculation = -4135
    }
    catch {
    }

    Update-ProgressWindow `
        -Text "Ouverture du classeur source..." `
        -Details $FichierSource

    $workbooks = $excel.Workbooks

    # Ouverture simplifiée :
    # 1 = chemin
    # 2 = mise à jour des liens : 0
    # 3 = lecture seule : $true
    $classeurSource = $workbooks.Open(
        $FichierSource,
        0,
        $true
    )

    if ($null -eq $classeurSource) {
        throw "Excel n'a pas réussi à ouvrir le classeur source."
    }

    $worksheets = $classeurSource.Worksheets
    $nombreOnglets = $worksheets.Count

    if ($nombreOnglets -eq 0) {
        throw "Le classeur ne contient aucun onglet."
    }

    $progressBar.Maximum = $nombreOnglets
    $progressBar.Value = 0

    # -----------------------------------------------------------------------
    # BOUCLE SUR LES ONGLETS
    # -----------------------------------------------------------------------

    for ($index = 1; $index -le $nombreOnglets; $index++) {

        $onglet = $null
        $nouveauClasseur = $null
        $nomOnglet = "Onglet inconnu"

        try {

            $onglet = $worksheets.Item($index)
            $nomOnglet = $onglet.Name

            Update-ProgressWindow `
                -Text "Préparation de l'onglet $index sur $nombreOnglets : $nomOnglet" `
                -Value ($index - 1) `
                -Details "Préparation du fichier de destination..."

            $nomFichier = Get-SafeFileName -Name $nomOnglet

            $cheminDestination = Get-UniqueFilePath `
                -Folder $DossierDestination `
                -FileName $nomFichier `
                -Extension ".xlsx"

            Update-ProgressWindow `
                -Text "Copie de l'onglet $index sur $nombreOnglets : $nomOnglet" `
                -Value ($index - 1) `
                -Details "Copie de la feuille dans un nouveau classeur..."

            try {
                $excel.CutCopyMode = 0
            }
            catch {
            }

            # Activation du classeur et de la feuille
            $classeurSource.Activate()
            $onglet.Activate()

            # Création d'un nouveau classeur contenant uniquement la feuille
            $onglet.Copy()

            $nouveauClasseur = $excel.ActiveWorkbook

            if ($null -eq $nouveauClasseur) {
                throw "Excel n'a pas créé le nouveau classeur."
            }

            try {
                $nouveauClasseur.CheckCompatibility = $false
            }
            catch {
            }

            Update-ProgressWindow `
                -Text "Enregistrement de l'onglet $index sur $nombreOnglets : $nomOnglet" `
                -Value ($index - 1) `
                -Details "Enregistrement du fichier : $nomFichier.xlsx"

            # 51 = xlOpenXMLWorkbook = XLSX
            # Appel simplifié pour éviter les problèmes COM
            $nouveauClasseur.SaveAs(
                $cheminDestination,
                51
            )

            Update-ProgressWindow `
                -Text "Fermeture du fichier $index sur $nombreOnglets : $nomOnglet" `
                -Value ($index - 1) `
                -Details $cheminDestination

            $nouveauClasseur.Close($false)

            Release-ComObject -ComObject $nouveauClasseur
            $nouveauClasseur = $null

            $nombreFichiersCrees++

            Update-ProgressWindow `
                -Text "Onglet $index sur $nombreOnglets terminé : $nomOnglet" `
                -Value $index `
                -Details "Fichier créé : $cheminDestination"
        }
        catch {

            $messageErreur = "Onglet '$nomOnglet' : $($_.Exception.Message)"
            $erreurs.Add($messageErreur)

            if ($null -ne $nouveauClasseur) {

                try {
                    $nouveauClasseur.Close($false)
                }
                catch {
                }

                Release-ComObject -ComObject $nouveauClasseur
                $nouveauClasseur = $null
            }
        }
        finally {

            if ($null -ne $onglet) {
                Release-ComObject -ComObject $onglet
                $onglet = $null
            }
        }
    }

    Update-ProgressWindow `
        -Text "Traitement terminé." `
        -Value $nombreOnglets `
        -Details "$nombreFichiersCrees fichier(s) créé(s)."

    $message = @"
Le traitement est terminé.

Nombre d'onglets détectés : $nombreOnglets
Nombre de fichiers créés : $nombreFichiersCrees

Dossier de destination :
$DossierDestination
"@

    if ($erreurs.Count -gt 0) {

        $message += "`n`nErreurs rencontrées :`n`n"
        $message += $erreurs -join "`n"

        [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Traitement terminé avec des erreurs",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
    else {

        $resultat = [System.Windows.Forms.MessageBox]::Show(
            "$message`n`nSouhaitez-vous ouvrir le dossier de destination ?",
            "Traitement terminé",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($resultat -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process explorer.exe -ArgumentList "`"$DossierDestination`""
        }
    }
}
catch {

    [System.Windows.Forms.MessageBox]::Show(
        "Une erreur générale est survenue :`n`n$($_.Exception.Message)",
        "Erreur",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
finally {

    # -----------------------------------------------------------------------
    # FERMETURE PROPRE D'EXCEL
    # -----------------------------------------------------------------------

    if ($null -ne $nouveauClasseur) {

        try {
            $nouveauClasseur.Close($false)
        }
        catch {
        }

        Release-ComObject -ComObject $nouveauClasseur
        $nouveauClasseur = $null
    }

    if ($null -ne $worksheets) {
        Release-ComObject -ComObject $worksheets
        $worksheets = $null
    }

    if ($null -ne $classeurSource) {

        try {
            $classeurSource.Close($false)
        }
        catch {
        }

        Release-ComObject -ComObject $classeurSource
        $classeurSource = $null
    }

    if ($null -ne $workbooks) {
        Release-ComObject -ComObject $workbooks
        $workbooks = $null
    }

    if ($null -ne $excel) {

        try {
            $excel.ScreenUpdating = $true
            $excel.EnableEvents = $true
        }
        catch {
        }

        try {
            $excel.Quit()
        }
        catch {
        }

        Release-ComObject -ComObject $excel
        $excel = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($null -ne $form) {

        try {
            $form.Close()
            $form.Dispose()
        }
        catch {
        }
    }
}