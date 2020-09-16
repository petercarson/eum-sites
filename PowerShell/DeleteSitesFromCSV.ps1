# Install-Module -Name ExchangeOnlineManagement

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    . $PSScriptRoot\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $Global:SPCredentials -Authentication Basic -AllowRedirection
  
#Import the session
Import-PSSession $Session -DisableNameChecking -AllowClobber

# -----------------------------------------
# 1. Look for any temp demo sites and delete them
# -----------------------------------------
# get all sites in the list that their Parent URL set to /sites/tempdemos

$connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL

$sitesListItems = Get-PnPListItem -Connection $connLandingSite -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <BeginsWith>
                <FieldRef Name='EUMDivision' />
                    <Value Type='Text'>Temporary Demos</Value>
            </BeginsWith>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMAlias'></FieldRef>
        <FieldRef Name='EUMSiteCreated'></FieldRef>
    </ViewFields>
</View>"

$sitesListItems | ForEach {
    $listItemID = $_["ID"]
    $listItemTitle = $_["Title"]
    $siteURL = $_["EUMSiteURL"]
    $alias = $_["EUMAlias"]
    $siteCreated = $_["EUMSiteCreated"]

    Write-Output "Remove site: $listItemTitle"
    
    if ($siteCreated -ne $null) {
        try {
            $connSite = Helper-Connect-PnPOnline -Url $siteURL
            $groupId = (Get-PnPSite -Includes GroupId).GroupId.Guid
            if ($groupId -ne "00000000-0000-0000-0000-000000000000") {
                #Delete the Office 365 Group
                Remove-UnifiedGroup -Identity $groupId -confirm:$False -ErrorAction Stop
                Start-Sleep -Seconds 30
                Remove-PnPTenantSite -Url $siteURL -Force -ErrorAction Stop -Connection $connSite
            }
            else {
                Remove-PnPTenantSite -Url $siteURL -Force -ErrorAction Stop -Connection $connSite
            }

            $connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL
            Remove-PnPListItem -List $SiteListName -Identity $listItemID -Force -Connection $connLandingSite
        }
        catch {
            Write-Verbose -Verbose -Message "Failed removing $($listItemTitle)..."
            Write-Verbose -Verbose -Message $_
        }
    }

}

#Remove the session
Remove-PSSession $Session
