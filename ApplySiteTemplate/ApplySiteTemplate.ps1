[string]$DistributionFolder = $Env:distributionFolder

if (-not $DistributionFolder)
{
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

[string]$SiteURL = Read-Host "Enter the URL of the site to apply the template to"

Helper-Connect-PnPOnline -Url $siteURL
$pnpSiteTemplate = $DistributionFolder + "\SiteTemplates\Client-Template-Template.xml"
Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate -ExcludeHandlers Publishing, ComposedLook, Navigation
Disconnect-PnPOnline
            
