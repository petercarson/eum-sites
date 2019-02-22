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

Set-PnPTraceLog -On -Level Debug
$pnpSiteTemplate = $DistributionFolder + "\SiteTemplates\Client-Template-Template.xml"
Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate

Add-PnPFolder -Name "Quotes" -Folder "/Shared Documents"
Add-PnPFolder -Name "Signed Quotes" -Folder "/Shared Documents/Quotes"
Add-PnPFolder -Name "Invoices" -Folder "/Shared Documents"

Add-PnPFolder -Name "Business Development" -Folder "/Private Documents"
Add-PnPFolder -Name "Confidential" -Folder "/Private Documents"
Add-PnPFolder -Name "Quotes" -Folder "/Private Documents"

Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"

$spWeb = Get-PnPWeb
$Title = $spWeb.Title
Remove-PnPGroup -Identity "$Title Members" -Force
Remove-PnPGroup -Identity "$Title Owners" -Force
Remove-PnPGroup -Identity "$Title Visitors" -Force

Disconnect-PnPOnline
            
