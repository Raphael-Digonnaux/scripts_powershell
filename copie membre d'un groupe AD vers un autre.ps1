# Remplacez ces variables avec les noms de vos groupes
$distributionGroup = "nom_groupe_source"
$securityGroup = "nom_groupe_cible"

# Récupérer les membres du groupe de distribution
$members = Get-ADGroupMember -Identity $distributionGroup

# Ajouter chaque membre au groupe de sécurité
foreach ($member in $members) {
    Add-ADGroupMember -Identity $securityGroup -Members $member
}
