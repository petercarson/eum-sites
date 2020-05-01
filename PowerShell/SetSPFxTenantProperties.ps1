Param
(
    [Parameter (Mandatory = $true)][string]$siteProvisioningApiUrl,
    [Parameter (Mandatory = $true)][string]$siteProvisioningApiClientID
)
$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

$connLandingSite = Helper-Connect-PnPOnline -Url $SitesListSiteURL

Set-PnPStorageEntity -Scope Tenant -Key siteProvisioningApiUrl -Value $siteProvisioningApiUrl -Comment "The Site Provisioning API URL" -Description "The Site Provisioning API URL" -Connection $connLandingSite
Set-PnPStorageEntity -Scope Tenant -Key siteProvisioningApiClientID -Value $siteProvisioningApiClientID -Comment "The Site Provisioning API Client ID" -Description "The Site Provisioning API Client ID" -Connection $connLandingSite