[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings -environmentId 1

ImportSiteColumns -ImportFile "$DistributionFolder\EUMSiteColumns.xml" -SiteUrl $SitesListSiteURL
ImportContentTypes -ImportFile "$DistributionFolder\EUMContentTypes.xml" -SiteUrl $SitesListSiteURL
