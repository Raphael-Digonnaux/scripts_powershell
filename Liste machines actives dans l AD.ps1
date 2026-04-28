# ======================================================================================
# Liste des machines "actives" dans l'AD
# (dernière connexion <= 90 jours)
#
# Sortie Excel dans le même répertoire que le script
# Fichier : MachinesActives_yyyy-MM-dd.xlsx
# ======================================================================================

[CmdletBinding()]
param(
    [int]$DelaiJours = 90
)

Import-Module ActiveDirectory
Import-Module ImportExcel

# Récupère le chemin du script
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Nom du fichier de sortie
$Date = Get-Date -Format "yyyy-MM-dd"
$ExcelFile = Join-Path $ScriptPath "MachinesActives_$Date.xlsx"

# Récupération des PC
$MachinesActives = Get-ADComputer -Filter * -Properties Name,OperatingSystem,lastlogondate |
    Where-Object {
        $_.LastLogonDate -ge (Get-Date).AddDays(-$DelaiJours)
    } |
    Select-Object Name,OperatingSystem,LastLogonDate |
    Sort-Object LastLogonDate -Descending

# Export Excel
$MachinesActives | Export-Excel -Path $ExcelFile -AutoSize -FreezeTopRow -Title "Machines AD actives"

Write-Host "Fichier créé : $ExcelFile" -ForegroundColor Green
