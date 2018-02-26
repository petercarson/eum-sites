[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Write-Host "Connecting to "$SitesListSiteURL
Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials
Write-Host "Applying the EUM Sites Template to "$SitesListSiteURL
Apply-PnPProvisioningTemplate -Path "$DistributionFolder\EUMSites.DeployTemplate.xml"
Disconnect-PnPOnline
