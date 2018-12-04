    [parameter(Mandatory=$true)][string]$TemplateSiteRelativeURL,

[string]$DistributionFolder = $Env:distributionFolder

if ($DistributionFolder -eq "")
{
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    $DistributionFolderArray = $DistributionFolder.Split('\')
    $DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
    $DistributionFolder = $DistributionFolderArray -join "\"
}

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

# -----------------------------
# Connect to the template site
# -----------------------------
[string]$TemplateSiteURL = Read-Host "Enter the URL of the site to template"
Helper-Connect-PnPOnline -Url $TemplateSiteURL

# If a $TemplateLocalPath wasn't provided as a parameter, use the template site's title and the current local path
[Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb

[string]$spWebTitle = $spWeb.Title
[string]$siteTemplateTitle = "$($spWebTitle.Replace(' ', '-'))-Template.xml"
$TemplateLocalFolder = $DistributionFolder + "\SiteTemplates\" + $siteTemplateTitle

# -------------------------
# Create the site template
# -------------------------
# Get-PnPProvisioningTemplate -out $TemplateLocalFolder -Handlers Fields, Lists, SiteSecurity, ContentTypes
Get-PnPProvisioningTemplate -out $TemplateLocalFolder -ExcludeHandlers Pages, Publishing, ComposedLook, Navigation
