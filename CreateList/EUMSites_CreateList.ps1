[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
$DistributionFolderArray = $DistributionFolder.Split('\')
$DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
$DistributionFolder = $DistributionFolderArray -join "\"

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Write-Host "Connecting to "$SitesListSiteURL
Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials

Write-Host "Applying the EUM Sites Template to "$SitesListSiteURL
Apply-PnPProvisioningTemplate -Path "$DistributionFolder\CreateList\EUMSites.DeployTemplate.xml"
Remove-PnPContentTypeFromList -List "Sites" -ContentType "Item"

Disconnect-PnPOnline
