[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

CreateSites