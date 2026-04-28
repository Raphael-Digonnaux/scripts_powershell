# ----------------------------------------
# Forcer l'exécution en STA (robuste)
# ----------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {

    Write-Host "Relancement en STA..." -ForegroundColor Yellow

    $scriptPath = $MyInvocation.MyCommand.Path

    if (-not $scriptPath) {
        Write-Host "Impossible de déterminer le chemin du script." -ForegroundColor Red
        exit
    }

    Start-Process -FilePath "powershell.exe" `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$scriptPath`"" `
        -Verb RunAs

    exit
}
# ----------------------------------------
# Chargement UI
# ----------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------------------
# Fenêtre de saisie
# ----------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sélection boîte partagée"
$form.Size = New-Object System.Drawing.Size(400,160)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Adresse de la boîte partagée :"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,20)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(360,20)
$textBox.Location = New-Object System.Drawing.Point(10,50)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(210,85)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Annuler"
$cancelButton.Location = New-Object System.Drawing.Point(290,85)

$okButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez saisir une adresse valide.")
    } else {
        $form.Tag = $textBox.Text
        $form.Close()
    }
})

$cancelButton.Add_Click({
    $form.Tag = $null
    $form.Close()
})

$form.Controls.AddRange(@($label,$textBox,$okButton,$cancelButton))
$form.Topmost = $true
$form.Add_Shown({$textBox.Select()})
$form.ShowDialog() | Out-Null

$SharedMailbox = $form.Tag

if (-not $SharedMailbox) {
    Write-Host "Opération annulée." -ForegroundColor Red
    exit
}

Write-Host "Boîte sélectionnée : $SharedMailbox" -ForegroundColor Green

# ----------------------------------------
# Connexion Exchange Online
# ----------------------------------------
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline

Write-Host "Connexion OK" -ForegroundColor Cyan

# ----------------------------------------
# Récupération des permissions
# ----------------------------------------
Write-Host "Récupération des permissions..." -ForegroundColor Yellow

# Full Access
$fullAccess = Get-MailboxPermission -Identity $SharedMailbox |
    Where-Object { $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF" } |
    Select-Object @{Name="Mailbox";Expression={$SharedMailbox}},
                  @{Name="User";Expression={$_.User}},
                  @{Name="Permission";Expression={"FullAccess"}}

# Send As
$sendAs = Get-RecipientPermission -Identity $SharedMailbox |
    Where-Object { $_.AccessRights -contains "SendAs" } |
    Select-Object @{Name="Mailbox";Expression={$SharedMailbox}},
                  @{Name="User";Expression={$_.Trustee}},
                  @{Name="Permission";Expression={"SendAs"}}

# Send on Behalf
$sendOnBehalf = Get-Mailbox -Identity $SharedMailbox |
    Select-Object -ExpandProperty GrantSendOnBehalfTo |
    ForEach-Object {
        [PSCustomObject]@{
            Mailbox    = $SharedMailbox
            User       = $_
            Permission = "SendOnBehalf"
        }
    }

# Fusion
$results = $fullAccess + $sendAs + $sendOnBehalf

# ----------------------------------------
# Export Excel
# ----------------------------------------
$exportPath = Join-Path -Path $PSScriptRoot -ChildPath ("Permissions_" + ($SharedMailbox -replace '@','_') + ".xlsx")

if (-not $results -or $results.Count -eq 0) {
    Write-Host "Aucune permission trouvée." -ForegroundColor Red
} else {
    $results | Format-Table -AutoSize

    $results | Export-Excel `
        -Path $exportPath `
        -WorksheetName "Permissions" `
        -AutoSize `
        -TableName "PermissionsMailbox" `
        -BoldTopRow `
        -FreezeTopRow `
        -AutoFilter

    Write-Host "Export Excel : $exportPath" -ForegroundColor Green
}

# ----------------------------------------
# Déconnexion
# ----------------------------------------
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Déconnecté." -ForegroundColor Cyan