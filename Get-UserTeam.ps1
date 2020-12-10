Import-Module MicrosoftTeams

Connect-MicrosoftTeams

$table = new-object System.Collections.ArrayList

$userInfo = New-Object System.Collections.ArrayList

$xlsxPath = "$PSScriptroot\List of BU Pigment Users_071220.xlsx"

foreach($row in (Import-XLSX -Path $xlsxPath -RowStart 1)) {
    $table.Add($row) | Out-Null
}

$table.mail | Select -Unique | ForEach-Object -Process { $userInfo.Add((Get-Team -User $_)) }

Disconnect-MicrosoftTeams