# Spécifiez le nom du groupe Active Directory 
$groupName = "Acces_vpn"

# Récupérer les membres du groupe
$groupMembers = Get-ADGroupMember -Identity $groupName

# Créer un objet Excel
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Ajouter les titres des colonnes
$worksheet.Cells.Item(1, 1) = "Nom des Utilisateurs"
$worksheet.Cells.Item(1, 2) = "Identifiant de l'utilisateur"
$worksheet.Cells.Item(1, 3) = "Email de l'utilisateur"
$worksheet.Cells.Item(1, 4) = "Service"

# Parcourir les membres du groupe et ajouter leurs informations à Excel
$row = 2
foreach ($member in $groupMembers) {

    # Récupérer l'utilisateur complet pour obtenir ses propriétés supplémentaires
    $user = Get-ADUser -Identity $member.SamAccountName -Properties mail, Department

    $cellName = $worksheet.Cells.Item($row, 1)
    $cellName.Value = $member.Name

    $cellId = $worksheet.Cells.Item($row, 2)
    $cellId.Value = $member.SamAccountName

    $cellEmail = $worksheet.Cells.Item($row, 3)
    $cellEmail.Value = $user.mail

    $cellService = $worksheet.Cells.Item($row, 4)
    $cellService.Value = $user.Department

    $row++
}

# Ajuster automatiquement la largeur des colonnes
$worksheet.Columns.AutoFit() | Out-Null

# Sauvegarder le fichier Excel
$excel.Visible = $true
$workbook.SaveAs("C:\Users\r.digonnaux\OneDrive - AEDE\Bureau\Script PS en cours\fichier.xlsx")
$excel.Quit()