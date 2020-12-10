Import-Module MicrosoftTeams

if ($null -eq $Creds) {
  $Creds = Get-Credential
}

Connect-MicrosoftTeams -Credential $Creds
#generate array for custom object
$teamsinfo = @()
$appsInfo = @()
$teamschannels = @()
#get all teams from organisation
$teams = Get-Team 
#find members, owner, guest
foreach ($team in $teams) {
  $displayname = ($team.DisplayName)
  $Description = ($team.Description)
  $groupid = $team.GroupId
  $members = (Get-TeamUser -GroupId $groupid -Role Member).User
  $owner = (Get-TeamUser -GroupId $groupid -Role Owner).User
  $guests = (Get-TeamUser -GroupId $groupid -Role Guest).User
  
  #custom object for output
  $teamsinfo += [pscustomobject]@{
    TeamID      = $groupid
    TeamName    = $displayname
    Owner       = ("$owner")
    Members     = ("$members")
    Guests      = ("$guests")
    Description = ("$Description")
    AppNames    = $apps.DisplayName
  }

  $apps = (Get-TeamsAppInstallation -TeamId $team.GroupId)

  foreach ($app in $apps) {
    $appsInfo += [PSCustomObject]@{
      TeamID   = $groupid
      TeamName = $displayname
      AppName  = $app.DisplayName
      Version  = $app.Version
    }
  }

  $tabs = (Get-TeamChannel -GroupId $team.GroupId)

  foreach ($tab in $tabs) {
    $teamschannels += [PSCustomObject]@{
      TeamID   = $groupid
      TeamName = $displayname
      TabName  = $tab.DisplayName
    }
  }
}
#show teaminformation in OutGrid-View
$teamsinfo | Sort-Object DisplayName | Export-Csv -Path .\Teams-GroupMembers.csv -Delimiter ";"

$appsInfo | Sort-Object DisplayName | Export-Csv -Path .\Teams-AppInfo.csv -Delimiter ";"

$teamschannels | Sort-Object DisplayName | Export-Csv -Path .\Teams-Channels.csv -Delimiter ";"