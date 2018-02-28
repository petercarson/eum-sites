if ($Env:POSTMethod)
{
    # POST method: $req
    $requestBody = Get-Content $req -Raw | ConvertFrom-Json
    $ID = $requestBody.id
}

[string]$DistributionFolder = $Env:distributionFolder

if (-not $DistributionFolder)
{
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

if ($listItemID)
{
    # Get the specific Site Collection List item in master site for the site that needs to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <And>
                    <Eq>
                        <FieldRef Name='ID'/>
                        <Value Type='Integer'>$itemId</Value>
                    </Eq>
                    <IsNull>
                        <FieldRef Name='EUMSiteCreated'/>
                    </IsNull>
                </And>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMPublicGroup'></FieldRef>
            <FieldRef Name='EUMSetComposedLook'></FieldRef>
            <FieldRef Name='EUMBrandingDeploymentType'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"
}
else
{
    # Check the Site Collection List in master site for any sites that need to be created
    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $pendingSiteCollections = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <IsNull>
                    <FieldRef Name='EUMSiteCreated'/>
                </IsNull>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMPublicGroup'></FieldRef>
            <FieldRef Name='EUMSetComposedLook'></FieldRef>
            <FieldRef Name='EUMBrandingDeploymentType'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='Author'></FieldRef>
        </ViewFields>
    </View>"
}

if ($pendingSiteCollections.Count -gt 0)
{
    # Get the time zone of the master site
    $spWeb = Get-PnPWeb -Includes RegionalSettings.TimeZone
    [int]$timeZoneId = $spWeb.RegionalSettings.TimeZone.Id

    # Iterate through the pending sites. Create them if needed, and apply template
    $pendingSiteCollections | ForEach {
        $pendingSite = $_

        [string]$siteTitle = $pendingSite["Title"]
        [string]$alias = $pendingSite["EUMAlias"]
        if ($alias)
        {
            $siteURL = "$($WebAppURL)/sites/$alias"
        }
        else
        {
            [string]$siteURL = ($pendingSite["EUMSiteURL"]).Url
        }
        [string]$publicGroup = $pendingSite["EUMPublicGroup"]
        [string]$breadcrumbHTML = $pendingSite["EUMBreadcrumbHTML"]
        [string]$parentURL = $pendingSite["EUMParentURL"].Url

        [bool]$siteCollection = CheckIfSiteCollection -siteURL $siteURL

        [string]$eumSiteTemplate = $pendingSite["EUMSiteTemplate"]

        [string]$author = $pendingSite["Author"].Email

        $baseSiteTemplate = ""
        $baseSiteType = ""
        $pnpSiteTemplate = ""
        $siteCreated = $false

        switch ($eumSiteTemplate)
        {
            "Classic Team Site"
                {
                $baseSiteTemplate = "STS#0"
                $baseSiteType = ""
                }
            "Modern Communication Site"
                {
                $baseSiteTemplate = ""
                $baseSiteType = "CommunicationSite"
                }
            "Modern Team Site"
                {
                $baseSiteTemplate = ""
                $baseSiteType = "TeamSite"
                }
            "Modern Client Site"
                {
                $baseSiteTemplate = ""
                $baseSiteType = "TeamSite"
                $pnpSiteTemplate = $DistributionFolder + "\SiteTemplates\Client-Template-Template.xml"
                }
        }

        # Classic style sites
        if ($baseSiteTemplate)
        {
            # Create the site
            if ($siteCollection)
            {
                # Create site (if it exists, it will error but not modify the existing site)
                Write-Output "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..."
                New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $author -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force
            }
            else
            {
                # Connect to parent site
                Helper-Connect-PnPOnline -Url $parentURL

                # Create the subsite
                Write-Output "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..."

                [string]$subsiteURL = $siteURL.Replace($parentURL, "").Trim('/')
                New-PnPWeb -Title $siteTitle -Url $subsiteURL -Template $baseSiteTemplate

                Disconnect-PnPOnline
            }
            $siteCreated = $true

        }
        # Modern style sites
        else
        {
            # Create the site
            Write-Output "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..."
            switch ($baseSiteType)
            {
                "CommunicationSite"
                {
                    New-PnPSite -Type CommunicationSite -Title $siteTitle -Url $siteURL
                    $siteCreated = $true
                }
                "TeamSite"
                {
                    if ($publicGroup)
                    {
                        New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias -IsPublic
                    }
                    else
                    {
                        New-PnPSite -Type TeamSite -Title $siteTitle -Alias $alias
                    }
                    $siteCreated = $true
                }
            }

        }

        if ($siteCreated)
        {
            if ($pnpSiteTemplate)
            {
                Helper-Connect-PnPOnline -Url $siteURL
                Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ExcludeHandlers Publishing, ComposedLook, Navigation
                Disconnect-PnPOnline
            }
            
            # Reconnect to the master site and update the site collection list
            Helper-Connect-PnPOnline -Url $SitesListSiteURL

            # Set the breadcrumb HTML
            [string]$siteRelativeURL = $siteURL.Replace($($WebAppURL), "")
            [string]$parentRelativeURL = $parentURL.Replace($($WebAppURL), "")
            $parentBreadcrumbHTML = ""
            if ($parentRelativeURL)
            {
                $parentListItem = GetSiteEntry -siteRelativeURL $parentRelativeURL
                if ($parentListItem)
                {
                    [string]$parentBreadcrumbHTML = $parentListItem["EUMBreadcrumbHTML"]
                }
            }
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

            # Set the site created date, breadcrumb, and site URL
            [Microsoft.SharePoint.Client.ListItem]$spListItem = Set-PnPListItem -List $SiteListName -Identity $pendingSite.Id -Values @{ "EUMSiteCreated" = [System.DateTime]::Now; "EUMBreadcrumbHTML" = $breadcrumbHTML; "EUMSiteURL" = $siteRelativeURL }
        }

        # Reconnect to the master site for the next iteration
        Helper-Connect-PnPOnline -Url $SitesListSiteURL
    }
}
else
{
    Write-Output "No sites pending creation"
}

