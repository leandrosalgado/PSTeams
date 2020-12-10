## Add credential to store
#Add-PnPStoredCredential -Name "MyCredential" -Username your@username.com

## Connect to Teams PnP
Connect-PnPOnline -Scopes "Group.ReadWrite.All" -Credentials "DPloyDEV"

#Store variables
$results = @()
$allTeams = Get-PnPTeamsTeam

#Loop through each Team
foreach ($team in $allTeams) {
    $allChannels = Get-PnPTeamsChannel -Team $team.DisplayName
    
    #Loop through each Channel
    foreach ($channel in $allChannels) {
        $allTabs = Get-PnPTeamsTab -Team $team.DisplayName -Channel $channel
        
        #Loop through each Tab + get the info!
        foreach ($tab in $allTabs) {
            $results += [PSCustomObject][ordered]@{
                TeamID      = $team.GroupId
                Team        = $team.DisplayName
                Visibility  = $team.Visibility
                ChannelName = $channel.DisplayName
                tabName     = $tab.DisplayName
            }
        }
    }
}
$results | Export-Csv -Path .\Tabs.csv -NoTypeInformation -Delimiter ";"