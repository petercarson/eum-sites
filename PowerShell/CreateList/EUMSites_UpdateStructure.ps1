$siteURL = Read-Host "Enter the URL of the site to deploy to"

Write-Host "Connecting to "$siteURL
$connLanding = Connect-PnPOnline -Url $siteURL -Interactive

$siteCollections = Get-PnPListItem -Connection $connLanding -List $SiteListName -Query "
<View>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMPublicGroup'></FieldRef>
        <FieldRef Name='EUMSiteVisibility'></FieldRef>
    </ViewFields>
</View>"

$siteCollections | ForEach-Object {
    $site = $_

    Write-Verbose -Verbose -Message "Updating list entry $($site["EUMSiteURL"]) with ID $($site.Id). Please wait..."
    if ($site["EUMPublicGroup"]) {
        [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $site.Id -Values @{ "EUMSiteVisibility" = "Public" } -Connection $connLanding
    }
    else {
        [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $site.Id -Values @{ "EUMSiteVisibility" = "Private" } -Connection $connLanding
    }
}