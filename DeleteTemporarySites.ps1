$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

$loadGraphAPICredentials = $true
LoadEnvironmentSettings

# Check the Site Collection List in master site for any sites that need to be created
Helper-Connect-PnPOnline -Url $SitesListSiteURL

$temporarySiteCollections = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <BeginsWith>
				<FieldRef Name='EUMParentURL'/>
				<Value Type='Url'>/sites/tempdemos</Value>
            </BeginsWith>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMAlias'></FieldRef>
    </ViewFields>
</View>"


if ($temporarySiteCollections.Count -gt 0) {
    # Iterate through the pending sites. Create them if needed, and apply template
    $temporarySiteCollections | ForEach {
        $temporarySite = $_

        if ($temporarySite["EUMAlias"] -eq $null) {
            Write-Output "Deleting non-group site $($_["Title"]), URL:$($_["EUMSiteURL"].Url)"
            Remove-PnPTenantSite -Url $_["EUMSiteURL"].Url -Force
            Remove-PnPListItem -List $SiteListName -Identity $_.Id -Force
        }
        else {
            Write-Output "Deleting group site $($_["Title"]), $($_["EUMSiteURL"].Url)"
            Connect-PnPOnline -AppId $AADClientID -AppSecret $AADSecret -AADDomain $AADDomain
            Remove-PnPUnifiedGroup -Identity $($_["EUMSiteURL"].Url)
            Helper-Connect-PnPOnline -Url $SitesListSiteURL
            Remove-PnPListItem -List $SiteListName -Identity $_.Id -Force
        }
    }
}
else {
    Write-Output "No temporary sites to delete"
}
