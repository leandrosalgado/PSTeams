Install-Module -Name PSExcel
Import-Module MicrosoftTeams
Import-Module PSExcel

Connect-MicrosoftTeams

$table = New-Object System.Collections.ArrayList

$userInfo = @()

$scriptPath = $PSScriptRoot

if ($psISE)
{
    $scriptPath = Split-Path -Path $psISE.CurrentFile.FullPath        
}

$xlsxPath = "$scriptPath\List of BU Pigment Users_071220.xlsx"

foreach($row in (Import-XLSX -Path $xlsxPath -RowStart 1)) {
    $table.Add($row) | Out-Null
}

foreach($user in ($table.mail | Select -Unique)) {
    foreach($team in (Get-Team -User $user)) {
        $userInfo += [PSCustomObject]@{
          UserEmail   = $user
          TeamID      = $team.GroupID
          TeamName    = $team.DisplayName
          Description = $team.Description
        }
    }
}

Disconnect-MicrosoftTeams

$userInfo | Sort-Object UserEmail | Export-Csv -Path .\User-Teams.csv -Delimiter ";"

