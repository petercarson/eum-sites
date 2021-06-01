$siteURL = Read-Host "Enter the URL of the site to deploy to"

Write-Host "Connecting to $siteURL"
$connLanding = Connect-PnPOnline -Url $siteURL -Interactive

Write-Host "Applying the EUM Sites Template to "$siteURL
Invoke-PnPSiteTemplate -Path "$DistributionFolder\CreateList\EUMSites.DeployTemplate.xml" -Connection $connLanding

Write-Host "Uploading EUMSites.SiteMetadataList.xml to PnP Templates library" 
Add-PnPFile -Path "$DistributionFolder\CreateList\EUMSites.SiteMetadataList.xml" -Folder "PnPTemplates" -Connection $connLanding
Add-PnPFile -Path "$DistributionFolder\CreateList\EUMSites.SiteMetadataListOnly.xml" -Folder "PnPTemplates" -Connection $connLanding

Disconnect-PnPOnline
