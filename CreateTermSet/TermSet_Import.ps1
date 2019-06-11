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

Helper-Connect-PnPOnline -Url $WebAppURL

#Import-PnPTaxonomy -Path "$DistributionFolder\CreateTermSet\EUMSitesTaxonomyPrivateDocumentType.txt"
#Import-PnPTaxonomy -Path "$DistributionFolder\CreateTermSet\EUMSitesTaxonomyProjectDocumentType.txt"

Import-PnPTermGroupFromXml -Path "$DistributionFolder\CreateTermSet\EUMSitesTaxonomy.txt"
