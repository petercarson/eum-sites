$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Connect-PnPOnline -Url $SitesListSiteURL -Credentials $SPCredentials -CreateDrive
New-Item -Path $pnpTemplatePath -ItemType "directory" -Force | out-null
Copy-Item -Path "spo:.\pnptemplates\*" -Destination $pnpTemplatePath -Force

function ApplyTemplate()
{
    Param
    (
        [Parameter(Mandatory=$true)][string] $siteURL
    )

    Helper-Connect-PnPOnline -Url $siteURL

    # Set the site collection admin
    if ($SiteCollectionAdministrator -ne "")
    {
        Add-PnPSiteCollectionAdmin -Owners $SiteCollectionAdministrator
    }

    # Remove the SharePoint groups
    Get-PnPGroup | Remove-PnPGroup -Force

    Set-PnPTraceLog -On -Level Debug
    $pnpSiteTemplate = "$pnpTemplatePath\Client-Template.xml"
    Apply-PnPProvisioningTemplate -Path $pnpSiteTemplate

    Add-PnPFolder -Name "Quotes" -Folder "/Shared Documents"
    Add-PnPFolder -Name "Signed Quotes" -Folder "/Shared Documents/Quotes"
    Add-PnPFolder -Name "Invoices" -Folder "/Shared Documents"

    Add-PnPFolder -Name "Business Development" -Folder "/Private Documents"
    Add-PnPFolder -Name "Confidential" -Folder "/Private Documents"
    Add-PnPFolder -Name "Quotes" -Folder "/Private Documents"

    Remove-PnPContentTypeFromList -List "Shared Documents" -ContentType "Document"
    Remove-PnPContentTypeFromList -List "Private Documents" -ContentType "Document"

    Disconnect-PnPOnline
}

$siteURL = Read-Host "Enter the URL of the site to apply the template to"
ApplyTemplate -siteURL $siteURL
