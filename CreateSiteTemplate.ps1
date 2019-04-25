$Global:AzureAutomation = (Get-Command "Get-AutomationVariable" -ErrorAction SilentlyContinue)
if ($AzureAutomation) { 
    . .\EUMSites_Helper.ps1
}
else {
    $DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
    . $DistributionFolder\EUMSites_Helper.ps1
}

LoadEnvironmentSettings

# -----------------------------
# Connect to the template site
# -----------------------------
[string]$TemplateSiteURL = Read-Host "Enter the URL of the site to template"
Helper-Connect-PnPOnline -Url $TemplateSiteURL

# Use the template site's title and the current local path
[Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb

[string]$spWebTitle = $spWeb.Title
[string]$siteTemplateTitle = "$($spWebTitle.Replace(' ', '-')).xml"
$TemplateFilename = "$pnpTemplatePath\$siteTemplateTitle"

# -------------------------
# Create the site template
# -------------------------
Get-PnPProvisioningTemplate -out $TemplateFilename -Handlers Fields, Lists, SiteSecurity, ContentTypes, PageContents
