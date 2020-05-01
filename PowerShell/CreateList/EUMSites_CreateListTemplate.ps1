[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
$DistributionFolderArray = $DistributionFolder.Split('\')
$DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
$DistributionFolder = $DistributionFolderArray -join "\"

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

Write-Host "Connecting to "$SitesListSiteURL
$connLanding = Helper-Connect-PnPOnline -Url $SitesListSiteURL


Write-Host "Creating the EUM Sites Template from "$SitesListSiteURL
$template = Get-PnPProvisioningTemplate -OutputInstance -Handlers Fields, Lists, ContentTypes -ContentTypeGroups "EUM Content Types" -ListsToExtract "Sites", "Divisions", "Site Templates", "Teams Channels", "Channel Templates", "Planner Templates", "PnP Templates", "Blacklisted Words" -Connection $connLanding
$template.BaseSiteTemplate = $null
$template.Lists | ForEach-Object { $_.Webhooks.Clear() }

# Remove all fields except those used by the content types
$fieldRefs = $template.ContentTypes | Select -ExpandProperty FieldRefs | Select -ExpandProperty Name
$fields = @() + $template.SiteFields # Deep copy the array of all fields
$template.SiteFields.Clear()
$fields | ForEach-Object {
    [xml]$schema = $_.SchemaXml
    $fieldName = $schema.Field.StaticName
    if ($fieldRefs.Contains($fieldName)) {
        $template.SiteFields.Add($_)
    }
}

Save-PnPProvisioningTemplate -InputInstance $template -Out "$DistributionFolder\CreateList\EUMSites.DeployTemplate.xml" -Force


$template = Get-PnPProvisioningTemplate -OutputInstance -Handlers Fields, Lists, ContentTypes -ContentTypeGroups "EUM Content Types" -ListsToExtract "Site Metadata" -Connection $connLanding
$template.BaseSiteTemplate = $null
$template.Lists | ForEach-Object { $_.Webhooks.Clear() }

$contentTypes = $template.ContentTypes | Where-Object { $_.Name -eq "Site Metadata" }
$template.ContentTypes.Clear()
$template.ContentTypes.Add($contentTypes)

# Remove all fields except those used by the content types
$fieldRefs = $contentTypes.FieldRefs | Select -ExpandProperty Name
$fields = @() + $template.SiteFields # Deep copy the array of all fields
$template.SiteFields.Clear()
$fields | ForEach-Object {
    [xml]$schema = $_.SchemaXml
    $fieldName = $schema.Field.StaticName
    if ($fieldRefs.Contains($fieldName)) {
        $template.SiteFields.Add($_)
    }
}

Save-PnPProvisioningTemplate -InputInstance $template -Out "$DistributionFolder\CreateList\EUMSites.SiteMetadataList.xml" -Force

$template = Get-PnPProvisioningTemplate -OutputInstance -Handlers Lists -ListsToExtract "Site Metadata" -Connection $connLanding
$template.BaseSiteTemplate = $null
$template.Scope = "Undefined"
$template.Lists | ForEach-Object { $_.Webhooks.Clear() }

Save-PnPProvisioningTemplate -InputInstance $template -Out "$DistributionFolder\CreateList\EUMSites.SiteMetadataListOnly.xml" -Force

Disconnect-PnPOnline