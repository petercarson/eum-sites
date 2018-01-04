[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings -environmentId 1

ImportLists -ImportFile "$DistributionFolder\EUMLists.xml" -SiteUrl $SitesListSiteURL
