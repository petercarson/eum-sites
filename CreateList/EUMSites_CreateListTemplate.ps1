[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
$DistributionFolderArray = $DistributionFolder.Split('\')
$DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
$DistributionFolder = $DistributionFolderArray -join "\"

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Write-Host "Connecting to "$SitesListSiteURL
Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials

Write-Host "Creating the EUM Sites Template from "$SitesListSiteURL
Get-PnPProvisioningTemplate -out "$DistributionFolder\CreateList\EUMSites.DeployTemplate.xml" -Handlers Fields, Lists, ContentTypes, PageContents

Disconnect-PnPOnline
