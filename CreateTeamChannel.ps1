Param
(
    [Parameter (Mandatory = $true)][int]$listItemID
)

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
    . .\Customizations.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
    . $DistributionFolder\Customizations.ps1
}

LoadEnvironmentSettings

try {
    Write-Verbose -Verbose -Message "Retrieving teams channel request details for listItemID $($listItemID)..."
    Helper-Connect-PnPOnline -Url $SitesListSiteURL
    $channelDetails = Get-PnPListItem -List $TeamsChannelsListName -Id $listItemID -Fields "ID", "Title", "IsPrivate", "Description", "TeamSiteURL", "Description", "CreateOneNoteSection", "CreateChannelPlanner", "ChannelTemplate"

    [string]$channelName = $channelDetails["Title"]
    [boolean]$isPrivate = $channelDetails["IsPrivate"]
    [string]$siteURL = $channelDetails["TeamSiteURL"]
    [string]$channelDescription = $channelDetails["Description"]
    [boolean]$createOneNote = $channelDetails["CreateOneNoteSection"]
    [boolean]$createPlanner = $channelDetails["CreateChannelPlanner"]
    [string]$channelTemplateId = $channelDetails["ChannelTemplate"].LookupId

    Disconnect-PnPOnline

    # Get the Office 365 Group ID
    Write-Verbose -Verbose -Message "Retrieving group ID for site $($siteURL)..."
    Helper-Connect-PnPOnline -Url $AdminURL
    $spSite = Get-PnPTenantSite -Url $siteURL
    $groupId = $spSite.GroupId
    Disconnect-PnPOnline
}
catch {
    Write-Error "Failed retrieving information for listItemID $($listItemID)"
    Write-Error $_
    exit    
}


try {
    # Create the new channel in Teams
    Write-Verbose -Verbose -Message "Creating channel $($channelName)..."
    $teamsConnection = Connect-MicrosoftTeams -Credential $SPCredentials
    $teamsChannel = New-TeamChannel -GroupId $groupId -DisplayName $channelName -Description $channelDescription
    $teamsChannelId = $teamsChannel.Id
    Disconnect-MicrosoftTeams

    if ($createOneNote) {
        Write-Verbose -Verbose -Message "Configuring OneNote for $($channelName)..."
        AddOneNoteTeamsChannelTab -groupId $groupId -channelName $channelName -teamsChannelId $teamsChannelId -siteURL $siteURL
    }

    if ($createPlanner) {
        Write-Verbose -Verbose -Message "Creating Planner for $($channelName)..."
        $planId = AddTeamPlanner -groupId $groupId -planTitle "$($channelName) Planner"
        AddPlannerTeamsChannelTab -groupId $groupId -planTitle "$($channelName) Planner" -planId $planId -channelName $channelName -teamsChannelId $teamsChannelId          
    }

    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    # Apply implementation specific customizations
    ApplyChannelCustomizations -listItemID $listItemID

    # update the SP list with the ChannelCreationDate
    Write-Verbose -Verbose -Message "Updating ChannelCreationDate..."

    $spListItem = Set-PnPListItem -List $TeamsChannelsListName -Identity $listItemID -Values @{"ChannelCreationDate" = (Get-Date) }
    Disconnect-PnPOnline
}
catch {
    Write-Error "Failed creating teams channel $($channelName)"
    Write-Error $_
    exit   
}