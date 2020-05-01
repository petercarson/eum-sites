function CheckSite() {
    Param
    (
        [parameter(Mandatory = $true)][string]$siteURL,
        [parameter(Mandatory = $false)][string]$parentURL
    )

    Write-Host "Checking to see if $SiteURL exists in the list..."

    $connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL

    $siteListItem = Get-PnPListItem -Connection $connLanding -List $SiteListName -Query "
    <View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='EUMSiteURL'/><Value Type='Text'>$SiteURL</Value>
                </Eq>
            </Where>
        </Query>
        <ViewFields>
            <FieldRef Name='ID'></FieldRef>
            <FieldRef Name='Title'></FieldRef>
            <FieldRef Name='EUMAlias'></FieldRef>
            <FieldRef Name='EUMDivision'></FieldRef>
            <FieldRef Name='EUMGroupSummary'></FieldRef>
            <FieldRef Name='EUMParentURL'></FieldRef>
            <FieldRef Name='SitePurpose'></FieldRef>
            <FieldRef Name='EUMSiteTemplate'></FieldRef>
            <FieldRef Name='EUMSiteURL'></FieldRef>
            <FieldRef Name='EUMSiteVisibility'></FieldRef>
            <FieldRef Name='EUMSiteCreated'></FieldRef>
            <FieldRef Name='EUMIsSubsite'></FieldRef>
            <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
        </ViewFields>
    </View>"
    
    if ($siteListItem.Count -eq 0) {
        Write-Host "Adding $siteURL to the list"
        $connSite = Helper-Connect-PnPOnline -Url $siteURL
        $Site = Get-PnPWeb  -Includes Created -Connection $connSite -ErrorAction Stop
        $siteCreatedDate = $Site.Created.Date
        if ($siteCreatedDate -eq $null) {
            $siteCreatedDate = Get-Date
        }

        [hashtable]$newListItemValues = @{ }

        $newListItemValues.Add("Title", $Site.Title)
        $newListItemValues.Add("EUMSiteURL", $SiteURL)
        $newListItemValues.Add("EUMSiteCreated", $siteCreatedDate)
        if ($parentURL -eq "") {
            $newListItemValues.Add("EUMIsSubsite", $false)
        }
        else {
            $newListItemValues.Add("EUMIsSubsite", $true)
            $newListItemValues.Add("EUMParentURL", $parentURL)
        }
        $newListItemValues.Add("EUMBreadcrumbHTML", (GetBreadcrumbHTML -siteURL $siteURL -siteTitle $Site.Title -parentURL $parentURL))

        $connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL
        [Microsoft.SharePoint.Client.ListItem]$newListItem = Add-PnPListItem -List $SiteListName -Values $newListItemValues -ContentType "EUM Site Collection List" -Connection $connLanding
    }
    else {
        [hashtable]$newListItemValues = @{ }

        $newListItemValues.Add("Title", $siteListItem["Title"])
        $newListItemValues.Add("EUMAlias", $siteListItem["EUMAlias"])
        $newListItemValues.Add("EUMDivision", $siteListItem["EUMDivision"].LookupValue)
        $newListItemValues.Add("EUMGroupSummary", $siteListItem["EUMGroupSummary"])
        $newListItemValues.Add("EUMParentURL", $siteListItem["EUMParentURL"])
        $newListItemValues.Add("SitePurpose", $siteListItem["SitePurpose"])
        $newListItemValues.Add("EUMSiteTemplate", $siteListItem["EUMSiteTemplate"])
        $newListItemValues.Add("EUMSiteURL", $siteListItem["EUMSiteURL"])
        $newListItemValues.Add("EUMSiteVisibility", $siteListItem["EUMSiteVisibility"])
        $newListItemValues.Add("EUMSiteCreated", $siteListItem["EUMSiteCreated"])
        $newListItemValues.Add("EUMIsSubsite", $siteListItem["EUMIsSubsite"])
        $newListItemValues.Add("EUMBreadcrumbHTML", $siteListItem["EUMBreadcrumbHTML"])
    }

    $connSite = Helper-Connect-PnPOnline -Url $siteURL

    #Check Site Metadata list exists
    $listSiteMetaData = "Site Metadata"
    Write-Host "Check if ""$($listSiteMetaData)"" list exists in $($siteURL) site. Updating..."
    $listExists = Get-PnPList -Identity $listSiteMetaData -Connection $connSite

    if (-not $listExists) {
        # Provision the Site Metadata list in the site and add the entry
        $retries = 0
        $pnpTemplateApplied = $false
        while (($retries -lt 20) -and ($pnpTemplateApplied -eq $false)) {
            try {
                $retries += 1
                if ($parentURL -eq "") {
                    $siteMetadataPnPTemplate = "$pnpTemplatePath\EUMSites.SiteMetadataList.xml"
                }
                else {
                    $siteMetadataPnPTemplate = "$pnpTemplatePath\EUMSites.SiteMetadataListOnly.xml"
                }

                Write-Verbose -Verbose -Message "Importing Site Metadata list with PnPTemplate $($siteMetadataPnPTemplate)"
                Apply-PnPProvisioningTemplate -Path $siteMetadataPnPTemplate -Connection $connSite -ErrorAction Stop
                Remove-PnPContentTypeFromList -List "Site Metadata" -ContentType "Item" -Connection $connSite

                [Microsoft.SharePoint.Client.ListItem]$spListItem = Add-PnPListItem -List "Site Metadata" -Values $newListItemValues -Connection $connSite
                $pnpTemplateApplied = $true
            }
            catch {      
                Write-Verbose -Verbose -Message "Failed applying PnP template."
                Write-Verbose -Verbose -Message $_.Exception | format-list -force
            }
        }
    }
    else {
        Write-Host "Updating Site Metadata List."
        Get-PnPListItem -List "Site Metadata" -Connection $connSite | Set-PnPListItem -List "Site Metadata" -Identity { $_.Id } -Values $newListItemValues -Connection $connSite
    }

    [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Webs -Connection $connSite

    if ($spWeb.Webs.Count -gt 0) {
        $spSubWebs = Get-PnPSubWebs -Web $spWeb -Connection $connSite
        foreach ($spSubWeb in $spSubWebs) {
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

# -----------------------------------------
# 1. Update any existing sites and delete all sites that no longer exist
# -----------------------------------------
# get all sites in the list that have Site Created set

$connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL

New-Item -Path $pnpTemplatePath -ItemType "directory" -Force | out-null
$pnpTemplates = Find-PnPFile -List "PnPTemplates" -Match *.xml -Connection $connLanding
$pnpTemplates | ForEach-Object {
    $File = Get-PnPFile -Url "$($SitesListSiteRelativeURL)/pnptemplates/$($_.Name)" -Path $pnpTemplatePath -AsFile -Force -Connection $connLanding
}

$siteCollectionListItems = Get-PnPListItem -Connection $connLanding -List $SiteListName -Query "
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
        <FieldRef Name='EUMAlias'></FieldRef>
        <FieldRef Name='EUMSiteVisibility'></FieldRef>
        <FieldRef Name='EUMBreadcrumbHTML'></FieldRef>
        <FieldRef Name='EUMParentURL'></FieldRef>
        <FieldRef Name='EUMSiteCreated'></FieldRef>
        <FieldRef Name='EUMSiteTemplate'></FieldRef>
        <FieldRef Name='EUMDivision'></FieldRef>
        <FieldRef Name='EUMCreateTeam'></FieldRef>
        <FieldRef Name='Author'></FieldRef>
        <FieldRef Name='EUMIsSubsite'></FieldRef>
    </ViewFields>
</View>"

Write-Output "Checking $($SiteListName) for updated and deleted sites. Please wait..."
$siteCollectionListItems | ForEach {
    $listItemID = $_["ID"]
    $listSiteTitle = $_["Title"]
    $alias = $_["EUMAlias"]
    $siteURL = $_["EUMSiteURL"]
    $parentURL = $_["EUMParentURL"]
    $division = $_["EUMDivision"].LookupValue
    $eumSiteTemplate = $_["EUMSiteTemplate"].LookupValue
    $listbreadcrumbHTML = $_["EUMBreadcrumbHTML"]
    $listSubsite = $_["EUMIsSubsite"]

    Write-Output "Checking if $listSiteTitle, URL:$siteURL still exists..."

    try {
        $connSite = Helper-Connect-PnPOnline -Url $siteURL
        $Site = Get-PnPWeb -Connection $connSite -ErrorAction Stop
        $siteExists = $true
    }
    catch [System.Net.WebException] {
        if ([int]$_.Exception.Response.StatusCode -eq 404) {
            $siteExists = $false
        }
        else {
            try {
                $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
                $spContext.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPCredentials.UserName, $SPCredentials.Password)
                $web = $spContext.Web
                $spContext.Load($web)
                $spContext.ExecuteQuery()
            }
            catch {      
                if (($_.Exception.Message -like "*Cannot contact site at the specified URL*") -and ($_.Exception.Message -like "*There is no Web named*")) {
                    $siteExists = $false
                }
            }
        }
    }

    if ($siteExists) {
        [string]$updatedBreadcrumbHTML = GetBreadcrumbHTML -siteURL $siteURL -siteTitle $Site.Title -parentURL $parentURL

        $charCount = ($siteURL.ToCharArray() | Where-Object { $_ -eq '/' } | Measure-Object).Count
        if ($charCount -le 4) {
            $updatedSubsite = $false
        }
        else {
            $updatedSubsite = $true
        }

        if (($listbreadcrumbHTML -notlike "*$($updatedBreadcrumbHTML)*") -or ($listSiteTitle -ne $Site.Title) -or ($listSubsite -ne $updatedSubsite)) {
            [hashtable]$newListItemValues = @{ }

            $newListItemValues.Add("Title", $Site.Title)
            $newListItemValues.Add("EUMBreadcrumbHTML", $updatedBreadcrumbHTML)
            $newListItemValues.Add("EUMIsSubsite", $updatedSubsite)

            Write-Host "$($Site.Title) exists in $($SiteListName) list. Updating..."
            $connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL
            [Microsoft.SharePoint.Client.ListItem]$newListItem = Set-PnPListItem -Identity $listItemID -List $SiteListName -Values $newListItemValues -Connection $connLanding
        }
        else {
            Write-Host "$($listSiteTitle) exists in $($SiteListName) list. No updates required."
        }
    }
    else {
        Write-Output "$listSiteTitle, URL:$siteURL does not exist. Deleting from list..."
        $connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL
        Remove-PnPListItem -List $SiteListName -Identity $listItemID -Force -Connection $connLanding
    }
}

# ---------------------------------------------------------
# 2. Iterate through all site collections and add new ones
# ---------------------------------------------------------
Write-Output "Adding tenant site collections to ($SiteListName). Please wait..."
$connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL
$siteCollections = Get-PnPTenantSite -IncludeOneDriveSites -Connection $connLanding

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
        ($SiteURL.ToLower() -ne "/")) {
        CheckSite -siteURL $SiteURL
    }
}
