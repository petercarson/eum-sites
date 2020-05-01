[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
$DistributionFolderArray = $DistributionFolder.Split('\')
$DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
$DistributionFolder = $DistributionFolderArray -join "\"

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Write-Host "Connecting to "$SitesListSiteURL
$connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL

Write-Host "Applying the EUM Sites Template to "$SitesListSiteURL
Apply-PnPProvisioningTemplate -Path "$DistributionFolder\CreateList\EUMSites.DeployTemplate.xml" -Connection $connLanding

Write-Host "Uploading EUMSites.SiteMetadataList.xml to PnP Templates library" 
Add-PnPFile -Path "$DistributionFolder\CreateList\EUMSites.SiteMetadataList.xml" -Folder "PnPTemplates" -Connection $connLanding
Add-PnPFile -Path "$DistributionFolder\CreateList\EUMSites.SiteMetadataListOnly.xml" -Folder "PnPTemplates" -Connection $connLanding

Disconnect-PnPOnline
