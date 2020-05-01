Param
(
    [Parameter (Mandatory = $false)][int]$listItemID
)

$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
. $DistributionFolder\EUMSites_Helper.ps1
. $DistributionFolder\CreateSite-Customizations.ps1

LoadEnvironmentSettings

$connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL
$pendingSiteCollections = Get-PnPListItem -Connection $connLandingSite -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <IsNull>
                        <FieldRef Name='EUMSiteCreated'/>
                    </IsNull>
                    <Eq>
                        <FieldRef Name='_ModerationStatus' />
                        <Value Type='ModStat'>0</Value>
                    </Eq>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
        </ViewFields>
    </View>"

$pendingSiteCollections | ForEach-Object {
    $pendingSite = $_
    $listItemID = $pendingSite["ID"]

    if (ProvisionSite -listItemID $listItemID) {
        # Apply and implementation specific customizations
        CreateSite-Customizations -listItemID $spListItem.Id
    }
}
