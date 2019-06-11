# ----------------------------------------------------------
# 
# Copyright Envision IT Inc. https://www.envisionit.com
# Licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License
# https://creativecommons.org/licenses/by-sa/3.0/deed.en_US
# 
# ----------------------------------------------------------

[string]$DistributionFolder = (Split-Path $MyInvocation.MyCommand.Path)
$DistributionFolderArray = $DistributionFolder.Split('\')
$DistributionFolderArray[$DistributionFolderArray.Count - 1] = ""
$DistributionFolder = $DistributionFolderArray -join "\"

. $DistributionFolder\EUMSites_Helper.ps1
LoadEnvironmentSettings

. $DistributionFolder\CreateTermSet\CoreBuilder_TermStore.ps1

$spContext = New-Object Microsoft.SharePoint.Client.ClientContext($WebAppURL)
$spContext.Credentials = $SPCredentials

ExportTermStore -PathToExportXMLTerms "$DistributionFolder\CreateTermSet" -XMLTermsFileName "EUMSitesTaxonomy.xml" -GroupToExport "EIT"

#Export-PnPTaxonomy -IncludeID -TermSet 376729aa-83dd-4ce7-b899-9ef06e674f67 -Path "$DistributionFolder\CreateTermSet\EUMSitesTaxonomyPrivateDocumentType.txt" -Force
#Export-PnPTaxonomy -IncludeID -TermSet 5d743221-2134-4277-8e1e-08455ea09703 -Path "$DistributionFolder\CreateTermSet\EUMSitesTaxonomyProjectDocumentType.txt" -Force

#Export-PnPTermGroupToXml -Out "$DistributionFolder\CreateTermSet\EUMSitesTaxonomy.txt" -Identity "e347a05e-222d-4bf4-90f7-e1b040b719df"
