@{
    RootModule = 'SharePointPnP.PowerShell.Online.Commands.dll'
    ModuleVersion = '3.2.1810.0'
    Description = 'SharePoint Patterns and Practices PowerShell Cmdlets for SharePoint Online'
    GUID = '8f1147be-a8e4-4bd2-a705-841d5334edc0'
    Author = 'SharePoint Patterns and Practices'
    CompanyName = 'SharePoint Patterns and Practices'
    DotNetFrameworkVersion = '4.5'
    ProcessorArchitecture = 'None'
    FunctionsToExport = '*'
    CmdletsToExport = 'Add-PnPApp','Add-PnPClientSidePage','Add-PnPClientSidePageSection','Add-PnPClientSideText','Add-PnPClientSideWebPart','Add-PnPContentType','Add-PnPContentTypeToDocumentSet','Add-PnPContentTypeToList','Add-PnPCustomAction','Add-PnPDataRowsToProvisioningTemplate','Add-PnPDocumentSet','Add-PnPEventReceiver','Add-PnPField','Add-PnPFieldFromXml','Add-PnPFieldToContentType','Add-PnPFile','Add-PnPFileToProvisioningTemplate','Add-PnPFolder','Add-PnPHtmlPublishingPageLayout','Add-PnPHubSiteAssociation','Add-PnPIndexedProperty','Add-PnPJavaScriptBlock','Add-PnPJavaScriptLink','Add-PnPListFoldersToProvisioningTemplate','Add-PnPListItem','Add-PnPMasterPage','Add-PnPNavigationNode','Test-PnPOffice365GroupAliasIsUsed','Add-PnPOffice365GroupToSite','Add-PnPProvisioningSequence','Add-PnPProvisioningSite','Add-PnPProvisioningSubSite','Add-PnPProvisioningTemplate','Add-PnPPublishingImageRendition','Add-PnPPublishingPage','Add-PnPPublishingPageLayout','Add-PnPRoleDefinition','Add-PnPSiteClassification','Add-PnPSiteCollectionAdmin','Add-PnPSiteCollectionAppCatalog','Add-PnPSiteDesign','Add-PnPSiteScript','Add-PnPStoredCredential','Add-PnPTaxonomyField','Add-PnPTenantCdnOrigin','Add-PnPTenantTheme','Add-PnPUserToGroup','Add-PnPView','Add-PnPWebhookSubscription','Add-PnPWebPartToWebPartPage','Add-PnPWebPartToWikiPage','Add-PnPWikiPage','Add-PnPWorkflowDefinition','Add-PnPWorkflowSubscription','Apply-PnPProvisioningHierarchy','Apply-PnPProvisioningTemplate','Set-PnPSitePolicy','Approve-PnPTenantServicePrincipalPermissionRequest','Clear-PnPListItemAsRecord','Clear-PnPRecycleBinItem','Clear-PnPTenantRecycleBinItem','Connect-PnPOnline','Connect-PnPMicrosoftGraph','Convert-PnPProvisioningTemplate','Convert-PnPFolderToProvisioningTemplate','Copy-PnPFile','Deny-PnPTenantServicePrincipalPermissionRequest','Disable-PnPFeature','Disable-PnPInPlaceRecordsManagementForSite','Disable-PnPPowerShellTelemetry','Disable-PnPResponsiveUI','Disable-PnPSiteClassification','Disable-PnPTenantServicePrincipal','Disconnect-PnPOnline','Enable-PnPFeature','Enable-PnPInPlaceRecordsManagementForSite','Enable-PnPPowerShellTelemetry','Enable-PnPResponsiveUI','Enable-PnPSiteClassification','Enable-PnPTenantServicePrincipal','Get-PnPProperty','Export-PnPTaxonomy','Export-PnPTermGroupToXml','Find-PnPFile','Get-PnPApp','Get-PnPAppInstance','Get-PnPAuditing','Get-PnPAuthenticationRealm','Get-PnPAvailableClientSideComponents','Get-PnPAzureADManifestKeyCredentials','Get-PnPClientSideComponent','Get-PnPClientSidePage','Get-PnPContentType','Get-PnPContentTypePublishingHubUrl','Get-PnPCustomAction','Get-PnPDefaultColumnValues','Get-PnPDocumentSetTemplate','Get-PnPEventReceiver','Get-PnPException','Get-PnPFeature','Get-PnPField','Get-PnPFile','Get-PnPFolder','Get-PnPFolderItem','Get-PnPGroup','Get-PnPGroupMembers','Get-PnPGroupPermissions','Get-PnPHealthScore','Get-PnPHideDefaultThemes','Get-PnPHomePage','Get-PnPHubSite','Get-PnPIndexedPropertyKeys','Get-PnPInPlaceRecordsManagement','Get-PnPJavaScriptLink','Get-PnPList','Get-PnPLabel','Get-PnPListInformationRightsManagement','Get-PnPListItem','Get-PnPListRecordDeclaration','Get-PnPMasterPage','Get-PnPNavigationNode','Get-PnPAccessToken','Get-PnPAzureCertificate','Get-PnPAppAuthAccessToken','Get-PnPConnection','Get-PnPSiteCollectionTermStore','Get-PnPStorageEntity','Get-PnPPowerShellTelemetryEnabled','Get-PnPPropertyBag','Get-PnPProvisioningSequence','Get-PnPProvisioningSite','Get-PnPProvisioningTemplate','Get-PnPProvisioningTemplateFromGallery','Get-PnPPublishingImageRendition','Get-PnPRecycleBinItem','Get-PnPRequestAccessEmails','Get-PnPRoleDefinition','Get-PnPSearchConfiguration','Get-PnPSearchCrawlLog','Get-PnPSite','Get-PnPSiteClassification','Get-PnPSiteClosure','Get-PnPSiteCollectionAdmin','Get-PnPSiteDesign','Get-PnPSiteDesignRights','Get-PnPSitePolicy','Get-PnPSiteScript','Get-PnPSiteSearchQueryResults','Get-PnPContext','Get-PnPStoredCredential','Get-PnPSubWebs','Get-PnPTaxonomyItem','Get-PnPTaxonomySession','Get-PnPTenant','Get-PnPTenantAppCatalogUrl','Get-PnPTenantCdnEnabled','Get-PnPTenantCdnOrigin','Get-PnPTenantCdnPolicies','Get-PnPTenantRecycleBinItem','Get-PnPTenantServicePrincipal','Get-PnPTenantServicePrincipalPermissionGrants','Get-PnPTenantServicePrincipalPermissionRequests','Get-PnPTenantSite','Get-PnPTenantTheme','Get-PnPTerm','Get-PnPTermGroup','Get-PnPTermSet','Get-PnPTheme','Get-PnPTimeZoneId','Get-PnPUnifiedGroup','Get-PnPUnifiedGroupMembers','Get-PnPUnifiedGroupOwners','Get-PnPUPABulkImportStatus','Get-PnPUser','Get-PnPUserProfileProperty','Get-PnPView','Get-PnPWeb','Get-PnPWebhookSubscriptions','Get-PnPWebPart','Get-PnPWebPartProperty','Get-PnPWebPartXml','Get-PnPWebTemplates','Get-PnPWikiPageContent','Get-PnPWorkflowDefinition','Get-PnPWorkflowInstance','Get-PnPWorkflowSubscription','Grant-PnPHubSiteRights','Grant-PnPSiteDesignRights','Grant-PnPTenantServicePrincipalPermission','Import-PnPAppPackage','Import-PnPTaxonomy','Import-PnPTermGroupFromXml','Import-PnPTermSet','Install-PnPApp','Install-PnPSolution','Invoke-PnPQuery','Invoke-PnPSiteDesign','Invoke-PnPWebAction','Measure-PnPList','Measure-PnPWeb','Measure-PnPResponseTime','Move-PnPClientSideComponent','Move-PnPFile','Move-PnPFolder','Move-PnPListItemToRecycleBin','Move-PnPRecycleBinItem','New-PnPExtensibilityHandlerObject','New-PnPGroup','New-PnPList','New-PnPPersonalSite','New-PnPAzureCertificate','New-PnPUnifiedGroup','New-PnPProvisioningCommunicationSite','New-PnPProvisioningHierarchy','New-PnPProvisioningSequence','New-PnPProvisioningTeamNoGroupSite','New-PnPProvisioningTeamNoGroupSubSite','New-PnPProvisioningTeamSite','New-PnPProvisioningTemplate','New-PnPProvisioningTemplateFromFolder','New-PnPSite','New-PnPTenantSite','New-PnPTerm','New-PnPTermGroup','New-PnPTermSet','New-PnPUPABulkImportJob','New-PnPUser','New-PnPWeb','Publish-PnPApp','Read-PnPProvisioningHierarchy','Read-PnPProvisioningTemplate','Register-PnPHubSite','Remove-PnPApp','Remove-PnPClientSideComponent','Remove-PnPClientSidePage','Remove-PnPContentType','Remove-PnPContentTypeFromDocumentSet','Remove-PnPContentTypeFromList','Remove-PnPCustomAction','Remove-PnPIndexedProperty','Remove-PnPEventReceiver','Remove-PnPField','Remove-PnPFieldFromContentType','Remove-PnPFile','Remove-PnPFileFromProvisioningTemplate','Remove-PnPFolder','Remove-PnPGroup','Remove-PnPHubSiteAssociation','Remove-PnPJavaScriptLink','Remove-PnPList','Remove-PnPListItem','Remove-PnPNavigationNode','Remove-PnPStorageEntity','Remove-PnPPropertyBagValue','Remove-PnPPublishingImageRendition','Remove-PnPRoleDefinition','Remove-PnPTenantSite','Remove-PnPSiteClassification','Remove-PnPSiteCollectionAdmin','Remove-PnPSiteCollectionAppCatalog','Remove-PnPSiteDesign','Remove-PnPSiteScript','Remove-PnPStoredCredential','Remove-PnPTaxonomyItem','Remove-PnPTenantCdnOrigin','Remove-PnPTenantTheme','Remove-PnPTermGroup','Remove-PnPUnifiedGroup','Remove-PnPUser','Remove-PnPUserFromGroup','Remove-PnPView','Remove-PnPWeb','Remove-PnPWebhookSubscription','Remove-PnPWebPart','Remove-PnPWikiPage','Remove-PnPWorkflowDefinition','Remove-PnPWorkflowSubscription','Rename-PnPFile','Rename-PnPFolder','Request-PnPReIndexList','Request-PnPReIndexWeb','Resolve-PnPFolder','Restore-PnPRecycleBinItem','Restore-PnPTenantRecycleBinItem','Resume-PnPWorkflowInstance','Revoke-PnPSiteDesignRights','Revoke-PnPTenantServicePrincipalPermission','Save-PnPProvisioningHierarchy','Save-PnPProvisioningTemplate','Send-PnPMail','Set-PnPAppSideLoading','Set-PnPAuditing','Set-PnPAvailablePageLayouts','Set-PnPClientSidePage','Set-PnPClientSideText','Set-PnPClientSideWebPart','Set-PnPContext','Set-PnPDefaultColumnValues','Set-PnPDefaultContentTypeToList','Set-PnPDefaultPageLayout','Set-PnPField','Set-PnPDocumentSetField','Set-PnPFileCheckedIn','Set-PnPFileCheckedOut','Set-PnPGroup','Set-PnPGroupPermissions','Set-PnPHideDefaultThemes','Set-PnPHomePage','Set-PnPHubSite','Set-PnPIndexedProperties','Set-PnPInPlaceRecordsManagement','Set-PnPList','Set-PnPLabel','Set-PnPListInformationRightsManagement','Set-PnPListItem','Set-PnPListItemAsRecord','Set-PnPListItemPermission','Set-PnPListPermission','Set-PnPListRecordDeclaration','Set-PnPMasterPage','Set-PnPMinimalDownloadStrategy','Set-PnPStorageEntity','Set-PnPPropertyBagValue','Set-PnPProvisioningTemplateMetadata','Set-PnPRequestAccessEmails','Set-PnPSearchConfiguration','Set-PnPSite','Set-PnPSiteClosure','Set-PnPSiteDesign','Set-PnPSiteScript','Set-PnPTaxonomyFieldValue','Set-PnPTenant','Set-PnPTenantCdnEnabled','Set-PnPTenantCdnPolicy','Set-PnPTenantSite','Set-PnPTheme','Set-PnPTraceLog','Set-PnPUnifiedGroup','Set-PnPUserProfileProperty','Set-PnPView','Set-PnPWeb','Set-PnPWebhookSubscription','Set-PnPWebPartProperty','Set-PnPWebPermission','Set-PnPWebTheme','Set-PnPWikiPageContent','Copy-PnPItemProxy','Move-PnPItemProxy','Start-PnPWorkflowInstance','Stop-PnPWorkflowInstance','Submit-PnPSearchQuery','Test-PnPListItemIsRecord','Test-PnPProvisioningHierarchy','Uninstall-PnPApp','Uninstall-PnPAppInstance','Uninstall-PnPSolution','Unpublish-PnPApp','Unregister-PnPHubSite','Update-PnPApp','Update-PnPSiteClassification'
    VariablesToExport = '*'
    AliasesToExport = '*'
    FormatsToProcess = 'SharePointPnP.PowerShell.Online.Commands.Format.ps1xml' 
    PrivateData = @{
        PSData = @{
            ProjectUri = 'https://aka.ms/sppnp'
            IconUri = 'https://raw.githubusercontent.com/pnp/media/master/optimized/pnp-projects/blue/png/pnp-powershell-300.png'
        }
    }
}
# SIG # Begin signature block
# MIIkVgYJKoZIhvcNAQcCoIIkRzCCJEMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBf05Q1L5umFRWP
# ewmJkDDt0ME8nL8llwZpXc0qSOBBMaCCDYEwggX/MIID56ADAgECAhMzAAABA14l
# HJkfox64AAAAAAEDMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMTgwNzEyMjAwODQ4WhcNMTkwNzI2MjAwODQ4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDRlHY25oarNv5p+UZ8i4hQy5Bwf7BVqSQdfjnnBZ8PrHuXss5zCvvUmyRcFrU5
# 3Rt+M2wR/Dsm85iqXVNrqsPsE7jS789Xf8xly69NLjKxVitONAeJ/mkhvT5E+94S
# nYW/fHaGfXKxdpth5opkTEbOttU6jHeTd2chnLZaBl5HhvU80QnKDT3NsumhUHjR
# hIjiATwi/K+WCMxdmcDt66VamJL1yEBOanOv3uN0etNfRpe84mcod5mswQ4xFo8A
# DwH+S15UD8rEZT8K46NG2/YsAzoZvmgFFpzmfzS/p4eNZTkmyWPU78XdvSX+/Sj0
# NIZ5rCrVXzCRO+QUauuxygQjAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUR77Ay+GmP/1l1jjyA123r3f3QP8w
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDM3OTY1MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAn/XJ
# Uw0/DSbsokTYDdGfY5YGSz8eXMUzo6TDbK8fwAG662XsnjMQD6esW9S9kGEX5zHn
# wya0rPUn00iThoj+EjWRZCLRay07qCwVlCnSN5bmNf8MzsgGFhaeJLHiOfluDnjY
# DBu2KWAndjQkm925l3XLATutghIWIoCJFYS7mFAgsBcmhkmvzn1FFUM0ls+BXBgs
# 1JPyZ6vic8g9o838Mh5gHOmwGzD7LLsHLpaEk0UoVFzNlv2g24HYtjDKQ7HzSMCy
# RhxdXnYqWJ/U7vL0+khMtWGLsIxB6aq4nZD0/2pCD7k+6Q7slPyNgLt44yOneFuy
# bR/5WcF9ttE5yXnggxxgCto9sNHtNr9FB+kbNm7lPTsFA6fUpyUSj+Z2oxOzRVpD
# MYLa2ISuubAfdfX2HX1RETcn6LU1hHH3V6qu+olxyZjSnlpkdr6Mw30VapHxFPTy
# 2TUxuNty+rR1yIibar+YRcdmstf/zpKQdeTr5obSyBvbJ8BblW9Jb1hdaSreU0v4
# 6Mp79mwV+QMZDxGFqk+av6pX3WDG9XEg9FGomsrp0es0Rz11+iLsVT9qGTlrEOla
# P470I3gwsvKmOMs1jaqYWSRAuDpnpAdfoP7YO0kT+wzh7Qttg1DO8H8+4NkI6Iwh
# SkHC3uuOW+4Dwx1ubuZUNWZncnwa6lL2IsRyP64wggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIWKzCCFicCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAQNeJRyZH6MeuAAAAAABAzAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgiEJ6KV+0
# xDNUWSrsOsBiu8ftT+wJqWQbsUQ+baqFa+QwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQCWIcXdSkEOFdKHknkl7MVv53cvR3a5rO36dsX3vs8K
# un+zDosePKkjpmRO0O61PCbriVPPeeviufek7eJZUJgq/0JJOOaeXM3TszkQFE/c
# bqCrU3Hz4TGtA9QkffkTyFhY2qrim6F5MIrdIDvPuYvfqLRriyITnrP9WWNaHaQe
# jFxjP4MhxPCT5mPq2ogO5yXVS3e/xDa0x2ajVMa5wfyLWM3i3UA50ycoEoxCNr5B
# wxLtjGWQuZRmBpjThUGzhFSsHS0Mhn9miQ0mm6oDIVuOS9akBm0lEPtsMnBiMbMp
# Z6ovpzNemO5ceLWS9M8s+hG9E+/hzmTCnolbutExazpDoYITtTCCE7EGCisGAQQB
# gjcDAwExghOhMIITnQYJKoZIhvcNAQcCoIITjjCCE4oCAQMxDzANBglghkgBZQME
# AgEFADCCAVcGCyqGSIb3DQEJEAEEoIIBRgSCAUIwggE+AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIGa2V9+SzfbOvy0U5rB1wJVyhE6CEAnpKpZEONBD
# rXJmAgZbrQQ9PAoYEjIwMTgxMDA5MDkwMTIyLjg1WjAHAgEBgAIB9KCB1KSB0TCB
# zjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEpMCcGA1UECxMg
# TWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xJjAkBgNVBAsTHVRoYWxl
# cyBUU1MgRVNOOjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBTZXJ2aWNloIIPHjCCBPUwggPdoAMCAQICEzMAAADUTxnD2ITLRWMA
# AAAAANQwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTAwHhcNMTgwODIzMjAyNjQwWhcNMTkxMTIzMjAyNjQwWjCBzjELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEpMCcGA1UECxMgTWljcm9zb2Z0IE9w
# ZXJhdGlvbnMgUHVlcnRvIFJpY28xJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjU4
# NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2
# aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAv0i7sMQTtT74OGpo
# uWATfTY+WnmWenhQyL9pJAFK/a7LZOF32Fmgf00gimwjJckiYY72FBQx0UqDb3aT
# susoplv47VDf27/0klDF49EJ8gBIwhEWLYTHtFNRg98M5wOTzVKfuhjWXK1nzPsW
# 5/Qx6NGoQjfKDhPjMEsRpEzYPH8v0ef742MSvrRI4ydNswRZRX0mcdrx2hRFmIRo
# Ku4m4jadyHxkwdYBWE4mA4V1vZfS0MxFHcjbCPaW1UfuHw4NnbHsDwRv8L0QOESu
# pwbQA7wledA0Xj+fQchPb69P0lk17eyRGKneUgB+fHAL0JNbY+qqMuwGh1R0Uda3
# yQEbJQIDAQABo4IBGzCCARcwHQYDVR0OBBYEFP5k/7/EzmU2uNamuNJbShVglJcH
# MB8GA1UdIwQYMBaAFNVjOlyKMZDzQ3t8RhvFM2hahW1VMFYGA1UdHwRPME0wS6BJ
# oEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01p
# Y1RpbVN0YVBDQV8yMDEwLTA3LTAxLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYB
# BQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljVGlt
# U3RhUENBXzIwMTAtMDctMDEuY3J0MAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYI
# KwYBBQUHAwgwDQYJKoZIhvcNAQELBQADggEBAEnI5BKJf8YlLqfOo0xeUA8dkzKY
# WUUpF1roTtBH5fmvsXLsVQx84IT8WKvAcqU2/2dnrMP/YFtz3qSrucNPZXIdjhEO
# WME9ELHrBKEWzHRmfm/DZuqorEPnQfaNboPXbJuyfxWc6tJn3h18adwiTJW7AIde
# Gf2e+D1v2qehOcrAtOC2l+rNRzCULRvKRgTS6o77geVG4V97yLKxLMF5ZdI89jN+
# q+8EHIsIB+ggteUYAGOX+WXZxR4/Ib/odk3ze9AD2FR+X1WlF1EzMrg9QUm4Aszm
# kld5wiR9HCbn7ji5EsAA8H6Irp8csHrRmXUl2WeVvysXxNjakP23cgQvEn8wggZx
# MIIEWaADAgECAgphCYEqAAAAAAACMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0xMDA3MDEyMTM2NTVa
# Fw0yNTA3MDEyMTQ2NTVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIIB
# IjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqR0NvHcRijog7PwTl/X6f2mU
# a3RUENWlCgCChfvtfGhLLF/Fw+Vhwna3PmYrW/AVUycEMR9BGxqVHc4JE458YTBZ
# sTBED/FgiIRUQwzXTbg4CLNC3ZOs1nMwVyaCo0UN0Or1R4HNvyRgMlhgRvJYR4Yy
# hB50YWeRX4FUsc+TTJLBxKZd0WETbijGGvmGgLvfYfxGwScdJGcSchohiq9LZIlQ
# YrFd/XcfPfBXday9ikJNQFHRD5wGPmd/9WbAA5ZEfu/QS/1u5ZrKsajyeioKMfDa
# TgaRtogINeh4HLDpmc085y9Euqf03GS9pAHBIAmTeM38vMDJRF1eFpwBBU8iTQID
# AQABo4IB5jCCAeIwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFNVjOlyKMZDz
# Q3t8RhvFM2hahW1VMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQ
# W9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNv
# bS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBa
# BggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MIGgBgNV
# HSABAf8EgZUwgZIwgY8GCSsGAQQBgjcuAzCBgTA9BggrBgEFBQcCARYxaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL1BLSS9kb2NzL0NQUy9kZWZhdWx0Lmh0bTBABggr
# BgEFBQcCAjA0HjIgHQBMAGUAZwBhAGwAXwBQAG8AbABpAGMAeQBfAFMAdABhAHQA
# ZQBtAGUAbgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAB+aIUQ3ixuCYP4FxAz2d
# o6Ehb7Prpsz1Mb7PBeKp/vpXbRkws8LFZslq3/Xn8Hi9x6ieJeP5vO1rVFcIK1GC
# RBL7uVOMzPRgEop2zEBAQZvcXBf/XPleFzWYJFZLdO9CEMivv3/Gf/I3fVo/HPKZ
# eUqRUgCvOA8X9S95gWXZqbVr5MfO9sp6AG9LMEQkIjzP7QOllo9ZKby2/QThcJ8y
# Sif9Va8v/rbljjO7Yl+a21dA6fHOmWaQjP9qYn/dxUoLkSbiOewZSnFjnXshbcOc
# o6I8+n99lmqQeKZt0uGc+R38ONiU9MalCpaGpL2eGq4EQoO4tYCbIjggtSXlZOz3
# 9L9+Y1klD3ouOVd2onGqBooPiRa6YacRy5rYDkeagMXQzafQ732D8OE7cQnfXXSY
# Ighh2rBQHm+98eEA3+cxB6STOvdlR3jo+KhIq/fecn5ha293qYHLpwmsObvsxsvY
# grRyzR30uIUBHoD7G4kqVDmyW9rIDVWZeodzOwjmmC3qjeAzLhIp9cAvVCch98is
# TtoouLGp25ayp0Kiyc8ZQU3ghvkqmqMRZjDTu3QyS99je/WZii8bxyGvWbWu3EQ8
# l1Bx16HSxVXjad5XwdHeMMD9zOZN+w2/XU/pnR4ZOC+8z1gFLu8NoFA12u8JJxzV
# s341Hgi62jbb01+P3nSISRKhggOsMIIClAIBATCB/qGB1KSB0TCBzjELMAkGA1UE
# BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
# BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEpMCcGA1UECxMgTWljcm9zb2Z0
# IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiUKAQEwCQYFKw4DAhoFAAMVAO0ICv9f7xxmgMPCjIj5XADfKo9DoIHe
# MIHbpIHYMIHVMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSkw
# JwYDVQQLEyBNaWNyb3NvZnQgT3BlcmF0aW9ucyBQdWVydG8gUmljbzEnMCUGA1UE
# CxMebkNpcGhlciBOVFMgRVNOOjU3RjYtQzFFMC01NTRDMSswKQYDVQQDEyJNaWNy
# b3NvZnQgVGltZSBTb3VyY2UgTWFzdGVyIENsb2NrMA0GCSqGSIb3DQEBBQUAAgUA
# 32azszAiGA8yMDE4MTAwOTEyNTcyM1oYDzIwMTgxMDEwMTI1NzIzWjBzMDkGCisG
# AQQBhFkKBAExKzApMAoCBQDfZrOzAgEAMAYCAQACAS8wBwIBAAICF+IwCgIFAN9o
# BTMCAQAwNgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAaAKMAgCAQACAxbj
# YKEKMAgCAQACAwehIDANBgkqhkiG9w0BAQUFAAOCAQEAMB/DC2ObL1HQfzpKyUCh
# vCQ2EqJPuOlKoR1EMA1A0Te51bKcSCwIP21WKnzoX2x4HGJIxHLXXreDAdZcOtOg
# KyIWxvA+o99aaLeiftYulJzsx5/P/z1gSv8V3RCaBHLcGnv6lsYVMm4flPZsU3zu
# tohnEFshYSeuCCev9pzCx1o9TfoFpAYJVc2OUM15mRhy2MK9PTCQKnNsEPy6YxIS
# og5JjnHnZwrmbkoMNQ1QajxdROkSGhHCpLo0aYmFU5Qup75L/OBLpdvgQL5JkgfV
# R0ojQW8oqKA7U3GlyEEKyJMuY2gO2cP5mixCfKZQA1yfiT7bf2vtaW5yoNkM6m1K
# vDGCAvUwggLxAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMz
# AAAA1E8Zw9iEy0VjAAAAAADUMA0GCWCGSAFlAwQCAQUAoIIBMjAaBgkqhkiG9w0B
# CQMxDQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIL9upE/hhHEC9tTMliRw
# KRhj7CSbQ4O2cIiiFTFigl54MIHiBgsqhkiG9w0BCRACDDGB0jCBzzCBzDCBsQQU
# 7QgK/1/vHGaAw8KMiPlcAN8qj0MwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMAITMwAAANRPGcPYhMtFYwAAAAAA1DAWBBQg8BwJNn6QEFVfDpN9
# ng1jVHvpIjANBgkqhkiG9w0BAQsFAASCAQAlo511euuWpvMDye7IC3V90pb+HOn9
# qKZhTSO4waS7jtNac4Vyyu2/inSIsxkqR2APKKYEdmEpxC9KqrifqWja3GqgiyF+
# C5B+oiLNLaVUgH2zbmS6yFbadqzqOlYGap9DoaSDYtReMx/zC2fxpn9dl0syV7Af
# q1Jl+1AiHwh50Xri4i0zfQIXoGAQXdU2N6HT2ALjssmfz3x/IS5NMoUY8IWnm8vN
# sqmd3uphVbBxEU2pNtEAyvXEn8auExfvJDKCOlHXgCN6n2NCHofq5SiYoASTcxe1
# +npiAdvL8Gw9UctIa1v1WeHG1YL0mJNAtsUFrBreZSaIN6rdJTBmGRa7
# SIG # End signature block
