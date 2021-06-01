function CheckSite() {
    Param
    (
        [parameter(Mandatory = $true)][string]$siteURL,
        [parameter(Mandatory = $false)][string]$parentURL
    )

    $connSite = Connect-PnPOnline -Url $siteURL

    #Check Site Metadata list exists
    $listSiteMetaData = "Site Metadata"
    Write-Host "Check if ""$($listSiteMetaData)"" list exists in $($siteURL) site. Updating..."
    $listExists = Get-PnPList -Identity $listSiteMetaData -Connection $connSite

    # Check if the list should be deleted and recreated
    if ($listExists) {
        Write-Host "Removing list..."
        Remove-PnPList -Identity "Site Metadata" -Force -Connection $connSite
    }

    [Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Includes Webs

    if ($spWeb.Webs.Count -gt 0) {
        $spSubWebs = Get-PnPSubWebs -Web $spWeb -Connection $connSite
        foreach ($spSubWeb in $spSubWebs) {
            CheckSite -siteURL $spSubWeb.Url -parentURL $siteURL
        }
    }
}

$siteURL = Read-Host "Enter a URL of a tenant site to connect to"

# ---------------------------------------------------------
# 1. Iterate through all site collections and subsites and remove list instances
# ---------------------------------------------------------
Write-Output "Checking all site collections..."
$connLanding = Connect-PnPOnline -Url $siteURL -Interactive
$siteCollections = Get-PnPTenantSite -Connection $connLanding

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
        ($SiteURL.ToLower() -notlike "https://envisionitdev.sharepoint.com/sites/EUMSitesTemplate")) {
        CheckSite -siteURL $SiteURL
    }
}

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
        ($SiteURL.ToLower() -ne $siteURL) -and 
        ($SiteURL.ToLower() -notlike "https://envisionitdev.sharepoint.com/sites/EUMSitesTemplate")) {
        Write-Host "Remove content type and fields from ""$($SiteURL)""..."
        $connSite = Connect-PnPOnline -Url $SiteURL

        $contentType = Get-PnPContentType -Identity "Site Metadata"
        if ($contentType.Count -eq 1) {
            Remove-PnPContentType -Identity "Site Metadata" -Force -Connection $connSite
        }

        if ($SiteURL -ne $siteURL) {
            Get-PnPField -Group "EUM Columns" -Connection $connSite | Remove-PnPField -Identity { $_.ID } -Force -Connection $connSite
        }
    }
}
