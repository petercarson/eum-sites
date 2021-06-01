# -----------------------------
# Connect to the template site
# -----------------------------
[string]$TemplateSiteURL = Read-Host "Enter the URL of the site to template"
$connLandingSite = Connect-PnPOnline -Url $TemplateSiteURL -Interactive

# Use the template site's title and the current local path
[Microsoft.SharePoint.Client.Web]$spWeb = Get-PnPWeb -Connection $connLandingSite

[string]$spWebTitle = $spWeb.Title
[string]$siteTemplateTitle = "$($spWebTitle.Replace(' ', '-')).xml"
$TemplateFilename = "$pnpTemplatePath\$siteTemplateTitle"

# -------------------------
# Create the site template
# -------------------------
Get-PnPSiteTemplate -out $TemplateFilename -Handlers Fields, Lists, ContentTypes, PageContents -Connection $connLandingSite

# -------------------------
# Upload the template back to SharePoint
# -------------------------
$UploadTemplate = Read-Host "Upload the template back to SharePoint (Y/N)?"
if (($UploadTemplate -eq "y") -or ($UploadTemplate -eq "Y")) {
    [string]$siteURL = Read-Host "Enter the URL of the landing site"
    $connLandingSite = Helper-Connect-PnPOnline -Url $siteURL
    $File = Add-PnPFile -Path $TemplateFilename -Folder "PnPTemplates" -Connection $connLandingSite
}
