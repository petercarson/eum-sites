function UpdateSiteEntry()
{
    Param
    (
        [parameter(Mandatory=$true)]$listItemID,
        [parameter(Mandatory=$true)][string]$siteURL,
        [parameter(Mandatory=$true)][string]$siteTitle,
        [parameter(Mandatory=$true)]$siteCreatedDate,
        [parameter(Mandatory=$true)]$parentURL,
        [parameter(Mandatory=$true)]$breadcrumbHTML
    )

    [string]$updatedBreadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

    $siteCollection = Get-PnPTenantSite -Url $siteURL
    $updatedTitle = $siteCollection.Title

    if (($breadcrumbHTML -notlike "*$($updatedBreadcrumbHTML)*") -or ($siteTitle -ne $updatedTitle))
    {
        [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $updatedTitle -breadcrumbHTML $updatedBreadcrumbHTML -siteCreatedDate $siteCreatedDate
        Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..."
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $listItemID -List $SiteListName -Values $newListItemValues
    }
    else
    {
        Write-Host "$($siteTitle) exists in $($SiteListName) list. No updates required."
    }

    # -----------
    # Subsites
    # -----------
    $spSubWebs = GetSubWebs -siteURL $SiteURL -disconnect
    if ($spSubWebs)
    {
        Write-Host "Checking subsites of $($siteTitle) to $($SiteListName). Please wait..."
        foreach ($spSubWeb in $spSubWebs)
        {
            [string]$siteURL = $spSubWeb.ServerRelativeUrl
            [string]$siteTitle = $spSubWeb.Title
            $siteCreatedDate = $spSubWeb.Created.Date
            [string]$breadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $siteTitle -parentURL $parentURL

            [hashtable]$newListItemValues = PrepareSiteItemValues -siteTitle $siteTitle -breadcrumbHTML $breadcrumbHTML -siteCreatedDate $siteCreatedDate

            $existingItem = GetSiteEntry -siteURL $siteURL -disconnect

            Helper-Connect-PnPOnline -Url $SitesListSiteURL
            if ($existingItem)
            {
                $updateRequired = $false

                foreach ($newListItemKey in $newListItemValues.Keys)
                {
                    if ($newListItemKey -eq "EUMBreadcrumbHTML")
                    {
                        if ($existingItem[$newListItemKey] -notlike "*$($newListItemValues[$newListItemKey])*")
                        {
                            $updateRequired = $true
                        }
                    }
                    elseif ($newListItemKey -ne "EUMSiteCreated")
                    {
                        if ($existingItem[$newListItemKey] -ne $newListItemValues[$newListItemKey])
                        {
                            $updateRequired = $true
                        }
                    }
                }

                if ($updateRequired)
                {
                    Write-Host "$($siteTitle) exists in $($SiteListName) list. Updating..."
                    [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $existingItem.Id -List $SiteListName -Values $newListItemValues
                }
                else
                {
                    Write-Host "$($siteTitle) exists in $($SiteListName) list. No updates required."
                }
            }
            else
            {
                [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "Base Site Request"
                Write-Host "The site $($siteTitle) was added to the $($SiteListName) list successfully"
            }
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
        <FieldRef Name='EUMSiteCreated'></FieldRef>
    </ViewFields>
</View>"

$siteCollectionListItems | ForEach {
    Write-Output "Checking if $($_["Title"]), URL:$($_["EUMSiteURL"]) needs updating..."
    $parentURL = $_["EUMParentURL"]
    if ($parentURL -eq $null)
    {
        $parentURL = ""
    }
	UpdateSiteEntry -listItemID $_["ID"] -SiteURL ($_["EUMSiteURL"]) -siteTitle $_["Title"] -siteCreatedDate $_["EUMSiteCreated"] -parentURL $parentURL -breadcrumbHTML $_["EUMBreadcrumbHTML"]
}
    
# ---------------------------------------------------------
# 3. Iterate through all site collections and add or update
# ---------------------------------------------------------
Write-Output "Adding tenant site collections to ($SiteListName). Please wait..."
Helper-Connect-PnPOnline -Url $SitesListSiteURL
$siteCollections = Get-PnPTenantSite -IncludeOneDriveSites

$siteCollections | ForEach {
    # Exclude the default site collections
    if (($SiteURL.ToLower() -notlike "*/portals/community") -and 
        ($SiteURL.ToLower() -notlike "*/portals/hub") -and 
        ($SiteURL.ToLower() -notlike "*/sites/contenttypehub") -and 
        ($SiteURL.ToLower() -notlike "*/search") -and 
        ($SiteURL.ToLower() -notlike "*/sites/appcatalog") -and 
        ($SiteURL.ToLower() -notlike "*/sites/compliancepolicycenter") -and 
        ($SiteURL.ToLower() -notlike "*-my.sharepoint.com*") -and 
        ($SiteURL.ToLower() -ne "/")) 
        {
            [string]$SiteURL = ($_["EUMSiteURL"]).Replace($WebAppURL, "")
            [string]$siteTitle = $_["Title"]
            [string]$siteCreated = $_["EUMSiteCreated"]

            [string]$SiteURL = ($_).Replace($WebAppURL, "")
            [string]$siteTitle = $_.Title
            [string]$breadcrumbHTML = GetBreadcrumbHTML -SiteURL $SiteURL -siteTitle $siteTitle -parentURL $_["EUMParentURL"]
            [string]$parentURL = ""

            [string]$parentBreadcrumbHTML = ""
            [string]$breadcrumbHTML = GetBreadcrumbHTML -SiteURL $SiteURL -siteTitle $siteTitle -parentBreadcrumbHTML $parentBreadcrumbHTML

            $spSubWebs = GetSubWebs -siteURL "$($WebAppURL)$($SiteURL)"
            Helper-Connect-PnPOnline -Url $_
            [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Created
            [DateTime]$siteCreatedDate = $spWeb.Created.Date

            [string]$SiteURL = ($_).Replace($WebAppURL, "")
            [string]$siteTitle = $_.Title
            Write-Output "Checking if $($_["Title"]), $($_["Url"]) needs to be added..."
	        AddSiteEntry -SiteURL $SiteURL -siteTitle $siteTitle -parentURL $parentURL -breadcrumbHTML $breadcrumbHTML -spSubWebs $spSubWebs -siteCreatedDate $siteCreatedDate    
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
