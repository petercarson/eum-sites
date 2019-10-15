$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
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

Helper-Connect-PnPOnline -Url $SitesListSiteURL

$sitesListItems = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='EUMParentURL' />
                    <Value Type='Text'>$Global:WebAppURL/sites/tempdemos</Value>
            </Eq>
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
        if ($alias -ne $null) {
            #Delete the Office 365 Group
            Remove-UnifiedGroup -Identity $alias -confirm:$False
        }
        else {
            Remove-PnPTenantSite -Url $siteURL -Force
        }
    }

    Remove-PnPListItem -List $SiteListName -Identity $listItemID -Force
}

#Remove the session
Remove-PSSession $Session
