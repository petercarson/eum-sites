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

# -------------------------------------------
# 2. Update existing entries
# -------------------------------------------
Write-Output "Updating existing entries in $($SiteListName). Please wait..."
Helper-Connect-PnPOnline -Url $SitesListSiteURL
$siteCollectionListItems = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <IsNotNull>
                <FieldRef Name='EUMSiteCreated'/>
            </IsNotNull>
        </Where>
        <OrderBy>
            <FieldRef Name='EUMParentURL' Ascending='TRUE' />
        </OrderBy>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
        <FieldRef Name='EUMParentURL'></FieldRef>
        <FieldRef Name='EUMSiteTemplate'></FieldRef>
        <FieldRef Name='EUMSiteCreated'></FieldRef>
    </ViewFields>
</View>"

$siteCollectionListItems | ForEach {
    [string]$SiteURL = $_["EUMSiteURL"]
    [string]$ParentURL = $_["EUMParentURL"]
    Write-Output "$SiteURL, $ParentURL"
}    

$siteCollectionListItems | ForEach {
    [string]$SiteRelativeURL = ($_["EUMSiteURL"]).Replace($WebAppURL, "")
    [string]$siteTitle = $_["Title"]
    [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL $_["EUMParentURL"]
    [string]$siteCreated = $_["EUMSiteCreated"]

    $spSubWebs = GetSubWebs -siteURL "$($WebAppURL)$($SiteRelativeURL)" -disconnect

    Write-Output "Checking if $($_["Title"]), URL:$($_["EUMSiteURL"]) needs updating..."
	AddOrUpdateSiteEntry -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreated -spSubWebs $spSubWebs    
}
    
# ---------------------------------------------------------
# 3. Iterate through all site collections and add or update
# ---------------------------------------------------------
Write-Output "Adding tenant site collections to ($SiteListName). Please wait..."
Helper-Connect-PnPOnline -Url $SitesListSiteURL
$siteCollections = Get-PnPTenantSite -IncludeOneDriveSites

$siteCollections | ForEach {
    # Exclude the default site collections
    if (($SiteRelativeURL.ToLower() -notlike "*/portals/community") -and 
        ($SiteRelativeURL.ToLower() -notlike "*/portals/hub") -and 
        ($SiteRelativeURL.ToLower() -notlike "*/sites/contenttypehub") -and 
        ($SiteRelativeURL.ToLower() -notlike "*/search") -and 
        ($SiteRelativeURL.ToLower() -notlike "*/sites/appcatalog") -and 
        ($SiteRelativeURL.ToLower() -notlike "*/sites/compliancepolicycenter") -and 
        ($SiteRelativeURL.ToLower() -notlike "*-my.sharepoint.com*") -and 
        ($SiteRelativeURL.ToLower() -ne "/")) 
        {
            [string]$SiteRelativeURL = ($_["EUMSiteURL"]).Replace($WebAppURL, "")
            [string]$siteTitle = $_["Title"]
            [string]$siteCreated = $_["EUMSiteCreated"]

            [string]$SiteRelativeURL = ($_).Replace($WebAppURL, "")
            [string]$siteTitle = $_.Title
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL $_["EUMParentURL"]
            [string]$parentURL = ""

            [string]$parentBreadcrumbHTML = ""
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

            $spSubWebs = GetSubWebs -siteURL "$($WebAppURL)$($SiteRelativeURL)"
            Helper-Connect-PnPOnline -Url $_
            [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Created
            [DateTime]$siteCreatedDate = $spWeb.Created.Date

            [string]$SiteRelativeURL = ($_).Replace($WebAppURL, "")
            [string]$siteTitle = $_.Title
            Write-Output "Checking if $($_["Title"]), $($_["Url"]) needs to be added..."
	        AddSiteEntry -siteRelativeURL $SiteRelativeURL -siteTitle $siteTitle -parentURL $parentURL -breadcrumbHTML $breadcrumbHTML -spSubWebs $spSubWebs -siteCreatedDate $siteCreatedDate    
        }
}

# -----------------------------------------
# 1. Delete all sites that no longer exist
# -----------------------------------------
# get all sites in the list that have Site Created set
$siteCollectionListItems = Get-PnPListItem -List $SiteListName -Query "
<View>
    <Query>
        <Where>
            <IsNotNull>
                <FieldRef Name='EUMSiteCreated'/>
            </IsNotNull>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID'></FieldRef>
        <FieldRef Name='Title'></FieldRef>
        <FieldRef Name='EUMSiteURL'></FieldRef>
        <FieldRef Name='EUMSetComposedLook'></FieldRef>
        <FieldRef Name='EUMBrandingDeploymentType'></FieldRef>
        <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
        <FieldRef Name='EUMParentURL'></FieldRef>
        <FieldRef Name='EUMSiteTemplate'></FieldRef>
    </ViewFields>
</View>"

Write-Output "Checking $($SiteListName) for deleted sites. Please wait..."
$siteCollectionListItems | ForEach {
    Write-Output "Checking if $($_["Title"]), URL:$($_["EUMSiteURL"]) still exists..."
    if (-not(CheckIfSiteExists -siteURL $_["EUMSiteURL"] -disconnect))
    {
        Write-Output "$($_["Title"]), URL:$($_["EUMSiteURL"]) does not exist. Deleting from list..."
        Helper-Connect-PnPOnline -Url $SitesListSiteURL
        Remove-PnPListItem -List $SiteListName -Identity $_.Id -Force
    }
}

