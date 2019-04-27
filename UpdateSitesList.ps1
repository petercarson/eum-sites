function CheckSite()
{
    Param
    (
    [parameter(Mandatory=$true)][string]$siteURL,
    [parameter(Mandatory=$false)][string]$parentURL
    )

    Write-Host "Checking to see if $SiteURL exists in the list..."

    Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $siteListItem = Get-PnPListItem -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='EUMSiteURL'/><Value Type='Text'>$SiteURL</Value>
                </Eq>
            </Where>
        </Query>
    </View>"
    
    if ($siteListItem.Count -eq 0)
    {
        Write-Host "Adding $siteURL to the list"
        Helper-Connect-PnPOnline -Url $siteURL
        $Site = Get-PnPWeb  -Includes Created -ErrorAction Stop
        $siteCreatedDate = $Site.Created.Date
        if ($siteCreatedDate -eq $null)
        {
            $siteCreatedDate = Get-Date
        }

        [string]$siteTitle = $Site.Title
        [string]$breadcrumbHTML = GetBreadcrumbHTML -SiteURL $SiteURL -siteTitle $siteTitle -parentURL $parentURL
        [hashtable]$newListItemValues = PrepareSiteItemValues -SiteURL $SiteURL -siteTitle $siteTitle -parentURL $parentURL -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreatedDate

        Helper-Connect-PnPOnline -Url $SitesListSiteURL
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List"
    }

    Helper-Connect-PnPOnline -Url $siteURL
    [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Webs

    if ($spWeb.Webs.Count -gt 0)
    {
        $spSubWebs = Get-PnPSubWebs -Web $spWeb
        foreach ($spSubWeb in $spSubWebs)
        {
            CheckSite -siteURL $spSubWeb.Url -parentURL $siteURL
        }
    }
}

$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

# ---------------------------------------------------------
# 2. Iterate through all site collections and add new ones
# ---------------------------------------------------------
Write-Output "Adding tenant site collections to ($SiteListName). Please wait..."
Helper-Connect-PnPOnline -Url $SitesListSiteURL
$siteCollections = Get-PnPTenantSite -IncludeOneDriveSites

$siteCollections | ForEach {
    [string]$SiteURL = $_.Url

    # Exclude the default site collections
    if (($SiteURL.ToLower() -notlike "*/portals/community") -and 
        ($SiteURL.ToLower() -notlike "*/portals/hub") -and 
        ($SiteURL.ToLower() -notlike "*/sites/contenttypehub") -and 
        ($SiteURL.ToLower() -notlike "*/search") -and 
        ($SiteURL.ToLower() -notlike "*/sites/appcatalog") -and 
        ($SiteURL.ToLower() -notlike "*/sites/compliancepolicycenter") -and 
        ($SiteURL.ToLower() -notlike "*-my.sharepoint.com*") -and 
        ($SiteURL.ToLower() -notlike "http://bot*") -and 
        ($SiteURL.ToLower() -ne "/")) 
        {
            CheckSite -siteURL $SiteURL
        }
}

# -----------------------------------------
# 1. Update any existing sites and delete all sites that no longer exist
# -----------------------------------------
# get all sites in the list that have Site Created set

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
        <FieldRef Name='EUMSiteCreated'></FieldRef>
    </ViewFields>
</View>"

Write-Output "Checking $($SiteListName) for updated and deleted sites. Please wait..."
$siteCollectionListItems | ForEach {
    $listItemID = $_["ID"]
    $siteURL = $_["EUMSiteURL"]
    $parentURL = $_["EUMParentURL"]
    $listSiteTitle = $_["Title"]
    $listbreadcrumbHTML = $_["EUMBreadcrumbHTML"]

    Write-Output "Checking if $listSiteTitle, URL:$siteURL still exists..."

    try
    {
        Helper-Connect-PnPOnline -Url $siteURL
        $Site = Get-PnPWeb -ErrorAction Stop
        $siteExists = $true
    }
    catch [System.Net.WebException]
    {
        if ([int]$_.Exception.Response.StatusCode -eq 404)
        {
            $siteExists = $false
        }
        else
        {
            try {
                $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                $spContext.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPCredentials.UserName, $SPCredentials.Password)
                $web = $spContext.Web
                $spContext.Load($web)
                $spContext.ExecuteQuery()
            }
            catch
            {      
                if (($_.Exception.Message -like "*Cannot contact site at the specified URL*") -and ($_.Exception.Message -like "*There is no Web named*"))
                {
                    $siteExists = $false
                }
            }
        }
    }

    if ($siteExists)
    {
        [string]$updatedBreadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $Site.Title -parentURL $parentURL
        if (($listbreadcrumbHTML -notlike "*$($updatedBreadcrumbHTML)*") -or ($listSiteTitle -ne $Site.Title))
        {
            [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $Site.Title -breadcrumbHTML $updatedBreadcrumbHTML
            Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..."
            [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $listItemID -List $SiteListName -Values $newListItemValues
        }
        else
        {
            Write-Host "$($listSiteTitle) exists in $($SiteListName) list. No updates required."
        }
    }
    else
    {
        Write-Output "$listSiteTitle, URL:$siteURL does not exist. Deleting from list..."
        Helper-Connect-PnPOnline -Url $SitesListSiteURL
        Remove-PnPListItem -List $SiteListName -Identity $listItemID -Force
    }
}

#        <Where>
#            <And>
#                <IsNotNull>
#                    <FieldRef Name='EUMSiteCreated'/>
#                </IsNotNull>
#                <Eq>
#                    <FieldRef Name='ID'/><Value Type='Number'>316</Value>
#                </Eq>
#            </And>
#        </Where>
