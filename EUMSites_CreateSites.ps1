try
{
    [string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

    . $DistributionFolder\EUMSites_Helper.ps1
    LoadEnvironmentSettings

    # Check the Site Collection List in master site for any sites that need to be created
    Connect-PnPOnline -Url $SitesListSiteURL -Credentials $credentials

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
        </ViewFields>
    </View>"

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
                    Write-Host "Creating site collection $($siteURL) with base template $($baseSiteTemplate). Please wait..." -ForegroundColor Yellow
                    New-PnPTenantSite -Title $siteTitle -Url $siteURL -Owner $AdminUserName -TimeZone $timeZoneId -Template $baseSiteTemplate -RemoveDeletedSite -Wait -Force
                }
                else
                {
                    # Connect to parent site
                    Connect-PnPOnline -Url $parentURL -Credentials $credentials

                    # Create the subsite
                    Write-Host "Creating subsite $($siteURL) with base template $($baseSiteTemplate) under $($parentURL). Please wait..." -ForegroundColor Yellow

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
                Write-Host "Creating site collection $($siteURL) with modern type $($baseSiteType). Please wait..." -ForegroundColor Yellow
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
                    Connect-PnPOnline -Url $siteURL -Credentials $credentials
                    Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ExcludeHandlers Publishing, ComposedLook, Navigation
                    Disconnect-PnPOnline
                }
            
                # Reconnect to the master site and update the site collection list
                Connect-PnPOnline -Url $SitesListSiteURL -Credentials $credentials

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
            Connect-PnPOnline -Url $SitesListSiteURL -Credentials $credentials
        }
    }
    else
    {
        Write-Host "No sites pending creation" -ForegroundColor Green
    }
}

catch [System.Management.Automation.CommandNotFoundException]
{
    Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Office 365 Dev PnP PowerShell CmdLets (https://github.com/SharePoint/PnP-PowerShell)"
    Write-Host "1. Install PowerShell Gallery from"
    Write-Host "`t https://www.powershellgallery.com/"
    Write-Host "2. Install PnP CmdLets. Execute the following PowerShell cmdlet:"
    Write-Host "`t Install-Module SharePointPnPPowerShellOnline"
    Write-Host "3. Install CredentialManager 2.0 (https://www.powershellgallery.com/packages/CredentialManager/2.0). Execute the following PowerShell cmdlet:"
    Write-Host "`t Install-Module -Name CredentialManager"
}
catch
{
    Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}