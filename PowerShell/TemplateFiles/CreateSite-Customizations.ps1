function CreateSite-Customizations {
    Param
    (
        [Parameter (Mandatory = $true)][int]$listItemID
    )

    $AzureAutomation = (Get-Command "Get-AutomationVariable" -errorAction SilentlyContinue)
    if ($AzureAutomation) {
        LoadEnvironmentSettings
    }

    Write-Verbose "CreateSite-Customizations Debug 1"

    $connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $pendingSiteCollection = Get-PnPListItem -Connection $connLandingSite -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='ID'/>
                    <Value Type='Integer'>$listItemID</Value>
                </Eq>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMSiteVisibility'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMCreateTeam'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"

    if ($pendingSiteCollection.Count -eq 1) {
        [string]$eumSiteTemplate = $pendingSiteCollection["EUMSiteTemplate"].LookupValue
        [string]$siteURL = $pendingSiteCollection["EUMSiteURL"]

        $siteTemplate = Get-PnPListItem -List "Site Templates" -Query "
                                                    <View>
                                                        <Query>
                                                            <Where>
                                                                <Eq>
                                                                    <FieldRef Name='Title'/>
                                                                    <Value Type='Text'>$eumSiteTemplate</Value>
                                                                </Eq>
                                                            </Where>
                                                        </Query>
                                                        <ViewFields>
                                                            <FieldRef Name='Title'></FieldRef>
                                                            <FieldRef Name='BaseClassicSiteTemplate'></FieldRef>
                                                            <FieldRef Name='BaseModernSiteType'></FieldRef>
                                                            <FieldRef Name='PnPSiteTemplate'></FieldRef>
                                                        </ViewFields>
                                                    </View>"
            
        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""

        if ($siteTemplate.Count -eq 1) {
            $baseSiteTemplate = $siteTemplate["BaseClassicSiteTemplate"]
            $baseSiteType = $siteTemplate["BaseModernSiteType"]
            if ($siteTemplate["PnPSiteTemplate"] -ne $null) {
                $pnpSiteTemplate = "$pnpTemplatePath\$($siteTemplate["PnPSiteTemplate"].LookupValue)"
            }
        }

        if ($siteTemplate -like "ABC") {
            Write-Verbose "Updating Site"
            Helper-Connect-PnPOnline -Url $siteURL
        }
        else {
            Write-Verbose "No customizations to apply"
        }

        # Reconnect to the master site and update the site collection list
        $connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL

        # Set the site created date
        [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $listItemID -Values @{ "EUMSiteCreated" = [System.DateTime]::Now } -Connection $connLandingSite
    }
}
