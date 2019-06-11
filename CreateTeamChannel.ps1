Param
(
    [Parameter (Mandatory = $true)][string]$RoleName,
    [Parameter (Mandatory = $true)][string]$RegistrationDisplayName,
    [Parameter (Mandatory = $true)][string]$siteURL
)

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

Helper-Connect-PnPOnline -Url $SitesListSiteURL
$queryURL = $siteURL.Replace($WebAppURL, "")
$siteDetails = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <Contains>
                <FieldRef Name='EUMSiteURL'/>
                <Value Type='URL'>$queryURL</Value>
            </Contains>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMSiteTemplate'></FieldRef>
    </ViewFields>
</View>"

if ($siteDetails["EUMSiteTemplate"] -eq "Modern Team Site")
{
    Write-Output "Creating Microsoft Teams Channel"

    # Pause the script to allow time for the modern site to finish provisioning
    Write-Output "Pausing for 120 seconds. Please wait..."
    Start-Sleep -Seconds 120
    Write-Output "Continuing..."

    # Get the Office 365 Group ID
    Helper-Connect-PnPOnline -Url $AdminURL
    $spSite = Get-PnPTenantSite -Url $siteURL
    $groupId = $spSite.GroupId
    Disconnect-PnPOnline

    # Create the new channel in Teams
    Connect-MicrosoftTeams -Credential $SPCredentials
    $channel = New-TeamChannel -GroupId $groupId -DisplayName $RegistrationDisplayName
    Disconnect-MicrosoftTeams

    # Create the corresponding folder in SharePoint and assign the appropriate rights
    Helper-Connect-PnPOnline -Url $siteURL
    Add-PnPFolder -Name $RegistrationDisplayName -Folder "Shared Documents"
    $folder = Get-PnPFolder -Url "Shared Documents/$RegistrationDisplayName" -Includes ListItemAllFields
    Set-PnPListItemPermission -List "Shared Documents" -Identity $folder.ListItemAllFields.Id -User $RoleName -AddRole "Read"
    Disconnect-PnPOnline
}
