$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials -CreateDrive

$siteCollectionListItems = Get-PnPListItem -List $SiteListName -Query "
<View>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMParentURL'></FieldRef>
    </ViewFields>
</View>"

$siteCollectionListItems | ForEach {
    $listItemID = $_["ID"]
    $listItemTitle = $_["Title"]
    $listItemSiteURL = $_["EUMSiteURL"].Url
    $listItemParentURL = $_["EUMParentURL"].Url

    Write-Host "Updating ID $listItemID - $listItemTitle"
    $Temp = Set-PnPListItem -Identity $listItemID -List $SiteListName -Values @{"EUMSiteURL2" = $listItemSiteURL; "EUMParentURL2" = $listItemParentURL}
}
