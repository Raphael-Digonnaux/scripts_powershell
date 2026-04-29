# ==============================================================================================
# SCRIPT : Audit des Permissions (Version Stabilisée)
# ==============================================================================================

# On s'assure que les erreurs s'affichent dans des boîtes de dialogue
$ErrorActionPreference = "Stop"

try {
    Add-Type -AssemblyName System.Windows.Forms
} catch {}

# --- 1. INTERFACE DE SAISIE ---
# On utilise une méthode plus simple qui ne nécessite pas de Runspace complexe
$inputBox = New-Object -ComObject MSScriptControl.ScriptControl
# Si l'UI complexe échoue, on peut utiliser Microsoft.VisualBasic pour une InputBox simple
Add-Type -AssemblyName Microsoft.VisualBasic
$SharedMailbox = [Microsoft.VisualBasic.Interaction]::InputBox("Entrez l'adresse de la boîte partagée :", "Sélection Boîte", "")

if ([string]::IsNullOrWhiteSpace($SharedMailbox)) { exit }

# --- 2. CONNEXION ET TRAITEMENT ---
try {
    # Import du module
    Import-Module ExchangeOnlineManagement
    
    # Connexion standard (La plus compatible)
    # Si vous n'utilisez aucun paramètre, le module ouvre une fenêtre de login standard.
    Connect-ExchangeOnline

    # Récupération
    $fullAccess = Get-MailboxPermission -Identity $SharedMailbox |
        Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF" } |
        Select-Object @{Name="Mailbox";Expression={$SharedMailbox}}, @{Name="User";Expression={$_.User}}, @{Name="Permission";Expression={"FullAccess"}}

    $sendAs = Get-RecipientPermission -Identity $SharedMailbox |
        Where-Object { $_.AccessRights -contains "SendAs" } |
        Select-Object @{Name="Mailbox";Expression={$SharedMailbox}}, @{Name="User";Expression={$_.Trustee}}, @{Name="Permission";Expression={"SendAs"}}

    $results = $fullAccess + $sendAs

    if ($results) {
        $currentDir = $PSScriptRoot
        if (-not $currentDir) { $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
        
        $fileName = "Permissions_" + ($SharedMailbox -replace '@','_') + ".xlsx"
        $exportPath = Join-Path -Path $currentDir -ChildPath $fileName

        # Export (nécessite ImportExcel sur le poste)
        $results | Export-Excel -Path $exportPath -WorksheetName "Permissions" -AutoSize -BoldTopRow -AutoFilter
        [System.Windows.Forms.MessageBox]::Show("Succès ! Fichier créé : $fileName")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Aucune permission trouvée.")
    }

} catch {
    [System.Windows.Forms.MessageBox]::Show("Erreur : $($_.Exception.Message)")
} finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}