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

# -----------------------------------------
# 1. Update any existing sites and delete all sites that no longer exist
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
        $Site = Get-PnPWeb
        $siteExists = $true
    }
    catch [System.Net.WebException]
    {
        $siteExists = $false
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
