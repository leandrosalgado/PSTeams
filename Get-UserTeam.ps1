Import-Module MicrosoftTeams

Connect-MicrosoftTeams

$table = new-object System.Collections.ArrayList

$userInfo = @()

$xlsxPath = "$PSScriptroot\List of BU Pigment Users_071220.xlsx"

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

