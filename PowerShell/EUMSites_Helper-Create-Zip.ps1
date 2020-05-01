$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

Copy-Item $DistributionFolder\EUMSites_Helper.ps1 -Destination $DistributionFolder\EUMSites_Helper.psm1 -Force
Compress-Archive -Path $DistributionFolder\EUMSites_Helper.psm1 -DestinationPath $DistributionFolder\EUMSites_Helper.zip -Force
Remove-Item $DistributionFolder\EUMSites_Helper.psm1

Copy-Item $DistributionFolder\CreateSite-Customizations.ps1 -Destination $DistributionFolder\CreateSite-Customizations.psm1 -Force
Compress-Archive -Path $DistributionFolder\CreateSite-Customizations.psm1 -DestinationPath $DistributionFolder\CreateSite-Customizations.zip -Force
Remove-Item $DistributionFolder\CreateSite-Customizations.psm1
